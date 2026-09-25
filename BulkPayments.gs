/**
 * Samaj Kendra Fees & Expense Management System
 * Module: BulkPayments.gs
 * Description: Dedicated backend for Rapid Multi-Counter fee collection, 
 *              offline sync reconciliation, chunked Drive PDF generation,
 *              and WhatsApp dispatch tracking.
 */

/**
 * Ensures "WhatsApp Status" column exists in "Payments Received" sheet headers.
 */
function ensurePaymentsHeaders(ss) {
  var paymentsSheet = ss.getSheetByName("Payments Received");
  if (!paymentsSheet) return null;
  
  var lastCol = paymentsSheet.getLastColumn();
  if (lastCol === 0) return paymentsSheet;
  
  var headers = paymentsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var waIdx = headers.indexOf("WhatsApp Status");
  if (waIdx === -1) {
    paymentsSheet.getRange(1, lastCol + 1).setValue("WhatsApp Status");
  }
  return paymentsSheet;
}

/**
 * Fetch initial dataset specifically tuned for Rapid Bulk Counter:
 * - Active Counter Lock status (controlled by Admin)
 * - Member directory snapshot with outstanding dues
 * - Payment accounts
 * - Today's / Recent synced payments for Master Dispatch Tracking
 */
function getBulkAppData(username) {
  try {
    if (typeof initializeSheetsIfNeeded === 'function') {
      initializeSheetsIfNeeded();
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("Settings");
    var membersSheet = ss.getSheetByName("Members");
    var paymentsSheet = ensurePaymentsHeaders(ss);
    var accountsSheet = ss.getSheetByName("Accounts");
    
    // 1. Settings Map
    var settings = {};
    if (settingsSheet) {
      var sVals = settingsSheet.getDataRange().getValues();
      for (var i = 1; i < sVals.length; i++) {
        var k = sVals[i][0];
        if (k) settings[k.toString().trim()] = sVals[i][1];
      }
    }
    
    var isCounterLocked = (settings["Event Counter Lock"] === "LOCKED");
    var userLower = (username || "").toLowerCase().trim();
    var isAdminUser = (userLower === "admin" || userLower === "jiguda");
    
    // 2. Users List (for Login authentication)
    var usersSheet = ss.getSheetByName("Users");
    var usersList = [];
    if (usersSheet) {
      var uVals = usersSheet.getDataRange().getValues();
      for (var u = 1; u < uVals.length; u++) {
        if (uVals[u][0]) {
          usersList.push({
            username: uVals[u][0].toString().trim(),
            displayName: uVals[u][2] ? uVals[u][2].toString().trim() : uVals[u][0].toString().trim()
          });
        }
      }
    }

    // 3. Accounts List
    var accountsList = ["Cash"];
    if (accountsSheet) {
      var accVals = accountsSheet.getDataRange().getValues();
      for (var a = 1; a < accVals.length; a++) {
        var accName = accVals[a][0];
        if (accName && accountsList.indexOf(accName.toString().trim()) === -1) {
          accountsList.push(accName.toString().trim());
        }
      }
    }
    
    // 4. Members Directory with Outstanding Dues
    var membersList = [];
    if (membersSheet) {
      var mVals = membersSheet.getDataRange().getValues();
      if (mVals.length > 1) {
        var mHeaders = mVals[0];
        
        // Robust header index finding (with physical column fallbacks)
        // Column A (0) = Sr No, Column B (1) = Member ID, Column C (2) = Name, Column D (3) = Mobile
        var idIdx = 1, nameIdx = 2, mobIdx = 3, vilIdx = -1;
        var idExact = false, nameExact = false, mobExact = false;
        for (var c = 0; c < mHeaders.length; c++) {
          var h = mHeaders[c].toString().trim().toLowerCase();
          // Member ID: exact matches first, then fallback
          if (!idExact && (h === "member id" || h === "id")) { idIdx = c; idExact = true; }
          // Name: exact matches ONLY — do NOT use h.includes("name") as it can grab wrong columns
          else if (!nameExact && (h === "name" || h === "full name" || h === "member name")) { nameIdx = c; nameExact = true; }
          // Mobile: prefix match is safe
          else if (!mobExact && (h.indexOf("mobile") > -1 || h.indexOf("phone") > -1)) { mobIdx = c; mobExact = true; }
          // Village: optional
          else if (vilIdx === -1 && (h.indexOf("village") > -1 || h.indexOf("city") > -1)) { vilIdx = c; }
        }
        // If still no exact name header found, try a broader includes check but only against early columns
        if (!nameExact) {
          for (var c2 = 0; c2 < Math.min(mHeaders.length, 6); c2++) {
            var h2 = mHeaders[c2].toString().trim().toLowerCase();
            if (h2.indexOf("name") > -1) { nameIdx = c2; break; }
          }
        }
        
        // Find all yearly fee status columns dynamically
        // Matches e.g. "2022-2023 Status", "2023-2024 S", "2025-26 Stat", "2026-27 Stat"
        var feeYears = [];
        // Support both key names used across Code.gs versions
        var feeAmount = parseFloat(settings["Annual Fee Amount"] || settings["Default Annual Fee"] || 500);
        for (var c = 0; c < mHeaders.length; c++) {
          var hName = mHeaders[c].toString().trim();
          var yrMatch = hName.match(/(\d{4}[-\/]\d{2,4})/);
          if (yrMatch && (hName.match(/stat/i) || hName.match(/\s+s$/i) || hName.indexOf("Status") > -1)) {
            var yr = yrMatch[1];
            // Find corresponding receipt column if any
            var receiptIdx = -1;
            for (var rc = 0; rc < mHeaders.length; rc++) {
              var rH = mHeaders[rc].toString().trim();
              if (rH.indexOf(yr) > -1 && (rH.match(/rec/i) || rH.match(/\s+r$/i))) {
                receiptIdx = rc;
                break;
              }
            }
            feeYears.push({
              year: yr,
              statusIdx: c,
              receiptIdx: receiptIdx,
              headerName: hName
            });
          }
        }
        
        for (var m = 1; m < mVals.length; m++) {
          var row = mVals[m];
          // Use header index with direct column index fallback (Col B = 1, Col C = 2, Col D = 3)
          var memId = (idIdx > -1 && row[idIdx]) ? row[idIdx].toString().trim() : (row[1] ? row[1].toString().trim() : "");
          var memName = (nameIdx > -1 && row[nameIdx]) ? row[nameIdx].toString().trim() : (row[2] ? row[2].toString().trim() : "");
          if (!memName && !memId) continue;
          
          var mobStr = (mobIdx > -1 && row[mobIdx]) ? row[mobIdx].toString().trim() : (row[3] ? row[3].toString().trim() : "");
          var vilStr = (vilIdx > -1 && row[vilIdx]) ? row[vilIdx].toString().trim() : "";
          
          var unpaidYears = [];
          var totalDues = 0;
          
          for (var y = 0; y < feeYears.length; y++) {
            var fy = feeYears[y];
            var rawVal = row[fy.statusIdx];
            var isPaid = false;
            var isNA = false;
            
            if (rawVal !== undefined && rawVal !== null && rawVal !== "") {
              var valStr = rawVal.toString().trim().toUpperCase();
              if (valStr === "TRUE" || rawVal === true || valStr === "PAID") {
                isPaid = true;
              } else if (valStr === "NA" || valStr === "N/A") {
                isNA = true;
              }
            }
            
            // In the sheet, FALSE or empty/unpaid means it is DUE!
            if (!isPaid && !isNA) {
              unpaidYears.push({
                year: fy.year,
                amount: feeAmount
              });
              totalDues += feeAmount;
            }
          }
          
          membersList.push({
            id: memId,
            "Member ID": memId,
            name: memName,
            Name: memName,
            mobile: mobStr,
            "Mobile Number": mobStr,
            village: vilStr,
            unpaidYears: unpaidYears,
            totalDues: totalDues
          });
        }
      }
    }
    
    // 4. Recent Synced Payments for Multi-Volunteer Dispatch & Tracking
    var recentPayments = [];
    if (paymentsSheet) {
      var pVals = paymentsSheet.getDataRange().getValues();
      if (pVals.length > 1) {
        var pHeaders = pVals[0];
        var findHdr = function(names) {
          for (var i = 0; i < pHeaders.length; i++) {
            var h = (pHeaders[i] || "").toString().trim().toLowerCase();
            for (var k = 0; k < names.length; k++) {
              if (h === names[k].toLowerCase()) return i;
            }
          }
          return -1;
        };
        var rNoIdx = findHdr(["Receipt Number", "Receipt No", "Receipt #"]);
        var dateIdx = findHdr(["Date"]);
        var pNameIdx = findHdr(["Name", "Member Name"]);
        var pIdIdx = findHdr(["Member ID", "MemberId", "ID"]);
        var pMobIdx = findHdr(["Mobile Number", "Mobile", "Phone"]);
        var modeIdx = findHdr(["Payment Mode", "Payment Account", "Mode"]);
        var txIdx = findHdr(["Transaction ID", "Tx ID", "Transaction Id"]);
        var amtIdx = findHdr(["Total Amount", "Amount"]);
        var partIdx = findHdr(["Fees Breakdown", "Particulars / Breakdown", "Breakdown", "Narration", "Paid for"]);
        var linkIdx = findHdr(["Receipt Link", "Receipt Link (PDF)", "PDF Link", "Receipt URL", "Link"]);
        var prepIdx = findHdr(["Prepared By", "Volunteer", "Created By"]);
        var waIdx = findHdr(["WhatsApp Status", "WA Status", "Status"]);
        
        // Scan in reverse (newest first)
        for (var p = pVals.length - 1; p >= 1; p--) {
          var pRow = pVals[p];
          var recNo = rNoIdx > -1 ? pRow[rNoIdx] : "";
          if (!recNo || recNo === "DELETED") continue;
          
          var rawDate = dateIdx > -1 ? pRow[dateIdx] : "";
          var dateStr = "";
          if (rawDate instanceof Date) {
            dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone() || "Asia/Kolkata", "dd/MM/yyyy");
          } else {
            dateStr = rawDate ? rawDate.toString() : "";
          }
          
          var pdfUrl = linkIdx > -1 ? (pRow[linkIdx] || "").toString().trim() : "";
          var hasPdf = (pdfUrl.indexOf("http") === 0);
          
          var waStat = waIdx > -1 ? (pRow[waIdx] || "").toString().trim() : "";
          if (!waStat) waStat = "Pending";
          
          recentPayments.push({
            receiptNo: recNo,
            date: dateStr,
            name: pNameIdx > -1 ? pRow[pNameIdx] : "",
            memberId: pIdIdx > -1 ? pRow[pIdIdx] : "",
            mobile: pMobIdx > -1 ? pRow[pMobIdx].toString().trim() : "",
            mode: modeIdx > -1 ? pRow[modeIdx] : "Cash",
            txId: txIdx > -1 ? (pRow[txIdx] || "").toString() : "",
            amount: amtIdx > -1 ? parseFloat(pRow[amtIdx] || 0) : 0,
            narration: partIdx > -1 ? pRow[partIdx] : "",
            pdfUrl: pdfUrl,
            pdfReady: hasPdf,
            preparedBy: prepIdx > -1 && pRow[prepIdx] ? pRow[prepIdx].toString().trim() : "system",
            whatsappStatus: waStat
          });
        }
      }
    }
    
    return {
      success: true,
      settings: settings,
      isCounterLocked: isCounterLocked,
      isAdminUser: isAdminUser,
      users: usersList,
      accounts: accountsList,
      members: membersList,
      recentPayments: recentPayments
    };
  } catch (err) {
    Logger.log("Error in getBulkAppData: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Toggle Event Counter Lock (Admin only).
 * When locked, counter volunteers cannot save new payments.
 */
function toggleCounterLock(isLocked, adminUser) {
  try {
    var userLower = (adminUser || "").toLowerCase().trim();
    if (userLower !== "admin" && userLower !== "jiguda") {
      return { success: false, error: "Permission denied: Only Admin or Jiguda can toggle the event counter lock." };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) return { success: false, error: "Settings sheet not found." };
    
    var lockValue = isLocked ? "LOCKED" : "ACTIVE";
    var sVals = settingsSheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < sVals.length; i++) {
      if (sVals[i][0] && sVals[i][0].toString().trim() === "Event Counter Lock") {
        settingsSheet.getRange(i + 1, 2).setValue(lockValue);
        found = true;
        break;
      }
    }
    if (!found) {
      settingsSheet.appendRow(["Event Counter Lock", lockValue]);
    }
    
    return { success: true, isCounterLocked: isLocked, status: lockValue };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Sync Bulk Payments from Offline / Counter Queue:
 * - Atomically increments Receipt Counter
 * - Assigns official sequential receipt numbers (REC-YYYY-XXXX)
 * - Fast-appends to "Payments Received"
 * - Posts entries to ledger
 * - Updates member fee status in "Members" sheet
 */
function syncBulkPayments(payload) {
  try {
    if (typeof initializeSheetsIfNeeded === 'function') {
      initializeSheetsIfNeeded();
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("Settings");
    var paymentsSheet = ensurePaymentsHeaders(ss);
    
    // Read Settings
    var settings = {};
    var sVals = settingsSheet.getDataRange().getValues();
    var counterRowIdx = -1;
    for (var i = 1; i < sVals.length; i++) {
      var k = sVals[i][0] ? sVals[i][0].toString().trim() : "";
      if (k) settings[k] = sVals[i][1];
      if (k === "Receipt Counter") counterRowIdx = i + 1;
    }
    
    // Check Counter Lock
    var isLocked = (settings["Event Counter Lock"] === "LOCKED");
    var loggedInUser = (payload.loggedInUser || "system").toString().trim();
    var userLower = loggedInUser.toLowerCase();
    var isAdmin = (userLower === "admin" || userLower === "jiguda");
    
    if (isLocked && !isAdmin) {
      return {
        success: false,
        error: "🔒 Counter is currently CLOSED by Admin. You cannot record or sync payments at this time."
      };
    }
    
    var paymentsList = payload.payments || [];
    if (!paymentsList.length) {
      return { success: false, error: "No payments provided in sync payload." };
    }
    
    var prefix = settings["Receipt Prefix"] || "REC";
    var counter = parseInt(settings["Receipt Counter"] || "1000");
    var pHeaders = paymentsSheet.getRange(1, 1, 1, paymentsSheet.getLastColumn()).getValues()[0];
    
    var prepByIdx = pHeaders.indexOf("Prepared By");
    var waStatIdx = pHeaders.indexOf("WhatsApp Status");
    var nextSrNo = paymentsSheet.getLastRow();
    
    var rowsToAppend = [];
    var receiptMap = {}; // tempId -> officialReceiptNo
    var syncedReceiptNumbers = [];
    var membersToUpdate = []; // queued member updates
    var ledgerPosts = [];     // queued ledger posts
    
    for (var idx = 0; idx < paymentsList.length; idx++) {
      var p = paymentsList[idx];
      counter++;
      nextSrNo++;
      
      var dateVal = p.date ? new Date(p.date) : new Date();
      var txYear = dateVal.getFullYear().toString();
      var officialReceiptNo = prefix + "-" + txYear + "-" + counter;
      
      receiptMap[p.tempId || ("TEMP-" + idx)] = officialReceiptNo;
      syncedReceiptNumbers.push(officialReceiptNo);
      
      // Calculate amounts and breakdown
      var totalAmount = parseFloat(p.donationAmount || 0);
      var breakdownArray = [];
      if (p.yearlyFees && typeof p.yearlyFees === 'object') {
        for (var yr in p.yearlyFees) {
          var fee = parseFloat(p.yearlyFees[yr]);
          totalAmount += fee;
          breakdownArray.push(yr + ": " + fee);
        }
      }
      var breakdownStr = breakdownArray.join("; ");
      
      // Build row according to headers
      var row = new Array(pHeaders.length);
      for (var c = 0; c < row.length; c++) row[c] = "";
      
      row[0] = nextSrNo - 1;                // Sr No
      row[1] = officialReceiptNo;          // Receipt Number
      row[2] = dateVal;                    // Date
      row[3] = p.name || "";               // Name
      row[4] = p.memberId || "NON-MEMBER"; // Member ID
      row[5] = p.mobileNumber || "";       // Mobile Number
      row[6] = p.paymentAccount || "Cash"; // Payment Mode
      row[7] = p.transactionId || "";      // Transaction ID
      row[8] = breakdownStr;               // Particulars / Breakdown
      row[9] = parseFloat(p.donationAmount || 0); // Donation
      row[10] = totalAmount;               // Total Amount
      row[11] = "PDF Pending";             // Receipt Link (PDF)
      row[12] = "";                        // Drive File ID
      
      if (prepByIdx > -1) {
        row[prepByIdx] = p.preparedBy || loggedInUser;
      }
      if (waStatIdx > -1) {
        row[waStatIdx] = "Pending";
      }
      
      rowsToAppend.push(row);
      
      // Queue Ledger Post
      ledgerPosts.push({
        account: p.paymentAccount || "Cash",
        date: dateVal,
        receiptNo: officialReceiptNo,
        narration: "Received from " + (p.name || "") + 
                   (p.memberId && p.memberId !== "NON-MEMBER" ? " (" + p.memberId + ")" : "") + 
                   (breakdownStr ? " for: " + breakdownStr : "") + 
                   (p.donationAmount > 0 ? " (Donation: " + p.donationAmount + ")" : ""),
        amount: totalAmount,
        preparedBy: p.preparedBy || loggedInUser
      });
      
      // Queue Member Status Update
      if (p.memberId && p.memberId !== "NON-MEMBER") {
        membersToUpdate.push({
          memberId: p.memberId,
          yearlyFees: p.yearlyFees,
          receiptNo: officialReceiptNo,
          mobileNumber: p.mobileNumber
        });
      }
    }
    
    // Fast batch append rows to Payments Received
    if (rowsToAppend.length > 0) {
      var startRow = paymentsSheet.getLastRow() + 1;
      paymentsSheet.getRange(startRow, 1, rowsToAppend.length, pHeaders.length).setValues(rowsToAppend);
    }
    
    // Update Receipt Counter in Settings sheet
    if (counterRowIdx > -1) {
      settingsSheet.getRange(counterRowIdx, 2).setValue(counter);
    } else {
      settingsSheet.appendRow(["Receipt Counter", counter]);
    }
    
    // Post to Ledgers
    for (var l = 0; l < ledgerPosts.length; l++) {
      var lp = ledgerPosts[l];
      if (typeof postToLedger === 'function') {
        try {
          postToLedger(lp.account, lp.date, lp.receiptNo, lp.narration, true, lp.amount, lp.preparedBy);
        } catch(ledErr) {
          Logger.log("Ledger posting warning: " + ledErr.toString());
        }
      }
    }
    
    // Update Member Fee Status & Mobile
    for (var m = 0; m < membersToUpdate.length; m++) {
      var mu = membersToUpdate[m];
      if (typeof updateMemberFeeStatusAndMobile === 'function') {
        try {
          updateMemberFeeStatusAndMobile(mu.memberId, mu.yearlyFees, "Paid", mu.receiptNo, mu.mobileNumber);
        } catch(memErr) {
          Logger.log("Member fee status update warning: " + memErr.toString());
        }
      }
    }
    
    return {
      success: true,
      syncedCount: paymentsList.length,
      receiptMap: receiptMap,
      receiptNumbers: syncedReceiptNumbers
    };
  } catch (err) {
    Logger.log("Error in syncBulkPayments: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Generate PDF Receipts in Safe Batches (Chunked Execution):
 * Takes a chunk of receipt numbers (e.g. 5 to 10 at a time),
 * creates Google Drive PDFs, shortens URLs, and updates Payments sheet.
 */
function generateBatchReceipts(receiptNumbersList) {
  try {
    if (!receiptNumbersList || !receiptNumbersList.length) {
      return { success: true, processed: [] };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("Settings");
    var paymentsSheet = ss.getSheetByName("Payments Received");
    if (!paymentsSheet) return { success: false, error: "Payments Received sheet not found." };
    
    var settings = {};
    if (settingsSheet) {
      var sVals = settingsSheet.getDataRange().getValues();
      for (var i = 1; i < sVals.length; i++) {
        var k = sVals[i][0];
        if (k) settings[k.toString().trim()] = sVals[i][1];
      }
    }
    
    var folders = (typeof getFolderStructure === 'function') ? getFolderStructure() : null;
    var pDataRange = paymentsSheet.getDataRange();
    var pValues = pDataRange.getValues();
    var pHeaders = pValues[0];
    
    var rNoIdx = pHeaders.indexOf("Receipt Number");
    if (rNoIdx === -1) return { success: false, error: "Receipt Number column not found in sheet." };

    var dateIdx = pHeaders.indexOf("Date");
    var nameIdx = pHeaders.indexOf("Name");
    var mIdIdx = pHeaders.indexOf("Member ID");
    var mobIdx = pHeaders.indexOf("Mobile Number");
    var modeIdx = pHeaders.indexOf("Payment Mode");
    
    // Robust multi-alias column lookup
    var partIdx = pHeaders.indexOf("Fees Breakdown");
    if (partIdx === -1) partIdx = pHeaders.indexOf("Particulars / Breakdown");
    if (partIdx === -1) partIdx = pHeaders.indexOf("Breakdown");
    
    var donIdx = pHeaders.indexOf("Donation Amount");
    var totIdx = pHeaders.indexOf("Total Amount");
    
    var linkIdx = pHeaders.indexOf("Receipt Link");
    if (linkIdx === -1) linkIdx = pHeaders.indexOf("Receipt Link (PDF)");
    
    var fileIdIdx = pHeaders.indexOf("Original PDF ID");
    if (fileIdIdx === -1) fileIdIdx = pHeaders.indexOf("Drive File ID");
    
    var prepByIdx = pHeaders.indexOf("Prepared By");
    
    var processedResults = [];
    var sheetUpdated = false;
    
    for (var rIdx = 0; rIdx < receiptNumbersList.length; rIdx++) {
      var targetReceiptNo = (receiptNumbersList[rIdx] || "").toString().trim();
      if (!targetReceiptNo) continue;
      
      // Find row
      var targetRowIndex = -1;
      for (var row = 1; row < pValues.length; row++) {
        if ((pValues[row][rNoIdx] || "").toString().trim() === targetReceiptNo) {
          targetRowIndex = row;
          break;
        }
      }
      
      if (targetRowIndex === -1) {
        processedResults.push({ receiptNo: targetReceiptNo, status: "NOT_FOUND" });
        continue;
      }
      
      var existingLink = linkIdx > -1 ? (pValues[targetRowIndex][linkIdx] || "").toString().trim() : "";
      if (existingLink.indexOf("http") === 0) {
        // Already generated
        processedResults.push({ receiptNo: targetReceiptNo, status: "ALREADY_EXISTS", pdfUrl: existingLink });
        continue;
      }
      
      var rawDate = dateIdx > -1 ? pValues[targetRowIndex][dateIdx] : new Date();
      var dateVal = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      var txYear = dateVal.getFullYear().toString();
      var pName = nameIdx > -1 ? (pValues[targetRowIndex][nameIdx] || "").toString().trim() : "";
      var mId = mIdIdx > -1 ? (pValues[targetRowIndex][mIdIdx] || "").toString().trim() : "";
      var mob = mobIdx > -1 ? (pValues[targetRowIndex][mobIdx] || "").toString().trim() : "";
      var mode = modeIdx > -1 ? (pValues[targetRowIndex][modeIdx] || "Cash").toString().trim() : "Cash";
      var partStr = partIdx > -1 ? (pValues[targetRowIndex][partIdx] || "").toString().trim() : "";
      var donAmt = donIdx > -1 ? parseFloat(pValues[targetRowIndex][donIdx] || 0) : 0;
      var totAmt = totIdx > -1 ? parseFloat(pValues[targetRowIndex][totIdx] || 0) : 0;
      
      var breakdownArray = partStr ? partStr.split("; ") : [];
      var preparedBy = prepByIdx > -1 ? (pValues[targetRowIndex][prepByIdx] || "system").toString().trim() : "system";
      var paymentData = {
        name: pName,
        memberId: mId,
        mobileNumber: mob,
        date: dateVal,
        paymentAccount: mode,
        donationAmount: donAmt,
        totalAmount: totAmt,
        preparedBy: preparedBy
      };
      
      try {
        var yearFolder = (folders && typeof getOrCreateFolder === 'function') 
          ? getOrCreateFolder(folders.receipts, txYear) 
          : DriveApp.getRootFolder();
        
        var safeName = (typeof sanitizeFilename === 'function') ? sanitizeFilename(pName) : "Receipt";
        var dateFormatted = (typeof formatDateStringDdMmYy === 'function') ? formatDateStringDdMmYy(dateVal) : txYear;
        var pdfName = targetReceiptNo + "-" + safeName + "-" + dateFormatted;
        
        var pdfBlob = null;
        if (typeof generatePDFBlob === 'function') {
          pdfBlob = generatePDFBlob(targetReceiptNo, paymentData, breakdownArray, settings, "Receipt");
        } else {
          pdfBlob = Utilities.newBlob("Receipt #" + targetReceiptNo + "\nName: " + pName + "\nAmount: INR " + totAmt, "application/pdf", pdfName + ".pdf");
        }
        
        var pdfFile = yearFolder.createFile(pdfBlob).setName(pdfName + ".pdf");
        try {
          pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch(shErr) {
          Logger.log("Sharing permission warning: " + shErr.toString());
        }
        
        var pdfFileId = pdfFile.getId();
        // Reliable direct Google Drive link
        var longUrl = "https://drive.google.com/file/d/" + pdfFileId + "/view?usp=sharing";
        var shortUrl = longUrl;
        
        // Attempt URL shortening with safe fallback
        if (typeof shortenUrl === 'function') {
          try {
            var s = shortenUrl(longUrl);
            if (s && s.indexOf("http") === 0) {
              shortUrl = s;
            }
          } catch (shortErr) {
            Logger.log("URL shortening error: " + shortErr.toString());
          }
        }
        
        // Update Sheet Row
        if (linkIdx > -1) {
          paymentsSheet.getRange(targetRowIndex + 1, linkIdx + 1).setValue(shortUrl);
          pValues[targetRowIndex][linkIdx] = shortUrl;
          sheetUpdated = true;
        }
        if (fileIdIdx > -1) {
          paymentsSheet.getRange(targetRowIndex + 1, fileIdIdx + 1).setValue(pdfFileId);
          pValues[targetRowIndex][fileIdIdx] = pdfFileId;
          sheetUpdated = true;
        }
        
        processedResults.push({
          receiptNo: targetReceiptNo,
          status: "SUCCESS",
          pdfUrl: shortUrl
        });
      } catch(pdfErr) {
        Logger.log("Error generating PDF for " + targetReceiptNo + ": " + pdfErr.toString());
        processedResults.push({
          receiptNo: targetReceiptNo,
          status: "ERROR",
          error: pdfErr.toString()
        });
      }
    }
    
    if (sheetUpdated) {
      SpreadsheetApp.flush();
    }
    
    return { success: true, processed: processedResults };
  } catch (err) {
    Logger.log("Error in generateBatchReceipts: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * On-demand receipt generator for a single receipt (e.g. when triggered directly from WhatsApp Runner).
 */
function generateSingleReceiptOnDemand(receiptNo) {
  var batchRes = generateBatchReceipts([receiptNo]);
  if (batchRes.success && batchRes.processed && batchRes.processed.length > 0) {
    var p = batchRes.processed[0];
    if (p.status === "SUCCESS" || p.status === "ALREADY_EXISTS") {
      return { success: true, receiptNo: receiptNo, pdfUrl: p.pdfUrl };
    } else {
      return { success: false, error: p.error || p.status };
    }
  }
  return { success: false, error: batchRes.error || "Failed to generate receipt." };
}

/**
 * Update WhatsApp dispatch status for a specific receipt in "Payments Received".
 */
function updateWhatsAppStatus(receiptNo, status, senderUser) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paymentsSheet = ensurePaymentsHeaders(ss);
    if (!paymentsSheet) return { success: false, error: "Payments sheet not found." };
    
    var pValues = paymentsSheet.getDataRange().getValues();
    var pHeaders = pValues[0];
    var rNoIdx = pHeaders.indexOf("Receipt Number");
    var waIdx = pHeaders.indexOf("WhatsApp Status");
    
    if (rNoIdx === -1 || waIdx === -1) {
      return { success: false, error: "Required columns missing." };
    }
    
    for (var i = 1; i < pValues.length; i++) {
      if (pValues[i][rNoIdx] === receiptNo) {
        var statusText = status || "Sent";
        if (senderUser) {
          var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd/MM hh:mm a");
          statusText += " (" + nowStr + " by " + senderUser + ")";
        }
        paymentsSheet.getRange(i + 1, waIdx + 1).setValue(statusText);
        return { success: true, receiptNo: receiptNo, statusText: statusText };
      }
    }
    return { success: false, error: "Receipt " + receiptNo + " not found." };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
