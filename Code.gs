/**
 * Membership & Expense Management System - Google Apps Script Backend (v3.3)
 * File: Code.js
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Member & Expense Manager')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Triggered automatically when the spreadsheet is opened.
 * Adds custom menu items to navigate to the HTML dashboard.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Member & Expense Manager')
    .addItem('Open Dashboard Sidebar', 'openDashboardSidebar')
    .addItem('Get Dashboard URL', 'openDashboardModal')
    .addItem('Re-Authorize Drive Access', 'authorizeDriveAccess')
    .addToUi();
}

/**
 * Run this function manually from Script Editor (or via menu) to authorize
 * Google Drive access. Required for PDF receipt/voucher generation.
 * Steps: Extensions > Apps Script > Select 'authorizeDriveAccess' > Run
 */
function authorizeDriveAccess() {
  var successMsg = 'Drive access is authorized successfully! PDF receipts will now work correctly.';
  
  try {
    var testFile = DriveApp.createFile('auth_test_delete_me.txt', 'auth test', MimeType.PLAIN_TEXT);
    testFile.setTrashed(true);
    
    Logger.log(successMsg);
    
    try {
      SpreadsheetApp.getUi().alert(successMsg);
    } catch(uiErr) {
      Logger.log('Note: Could not show popup dialog because the script is running in a non-UI context (e.g. from the Script Editor). The authorization was still successful!');
    }
  } catch(e) {
    var errorMsg = 'Drive authorization failed: ' + e.toString() + '\n\nPlease check your OAuth scopes in appsscript.json and redeploy.';
    Logger.log(errorMsg);
    
    try {
      SpreadsheetApp.getUi().alert(errorMsg);
    } catch(uiErr) {
      Logger.log('Note: Could not show popup dialog because the script is running in a non-UI context (e.g. from the Script Editor).');
    }
  }
}

function openDashboardSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Member & Expense Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openDashboardModal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = getSettingsSheet(ss);
  var settings = getSettingsMap(settingsSheet);
  var webAppUrl = settings["Web App Link"];
  
  if (!webAppUrl || webAppUrl.indexOf("http") === -1) {
    webAppUrl = ScriptApp.getService().getUrl();
  }
  
  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert("Please deploy this script as a Web App first (Deploy > New deployment) to get the dashboard link.");
    return;
  }
  
  // Save Web App link to Settings sheet if missing
  if (!settings["Web App Link"] || settings["Web App Link"] !== webAppUrl) {
    updateSettingValue("Web App Link", webAppUrl);
  }
  
  var htmlContent = `
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 20px; text-align: center; color: #1F2937; }
        .title { font-size: 16px; font-weight: bold; margin-bottom: 10px; }
        .btn { display: inline-block; padding: 10px 20px; background-color: #4F46E5; color: white; text-decoration: none; border-radius: 8px; font-weight: bold; margin-top: 15px; margin-bottom: 15px; font-size: 14px; }
        .btn:hover { background-color: #4338CA; }
        .url-box { font-size: 11px; color: #6B7280; word-break: break-all; background: #F3F4F6; padding: 10px; border-radius: 6px; border: 1px solid #E5E7EB; }
      </style>
    </head>
    <body>
      <div class="title">Dashboard Web Link</div>
      <p style="font-size: 13px;">Click below to open the manager dashboard in a new tab:</p>
      <a href="${webAppUrl}" target="_blank" class="btn" onclick="google.script.host.close()">Open Dashboard</a>
      <div class="url-box">${webAppUrl}</div>
      <p style="font-size: 11px; color: #9CA3AF; margin-top: 10px;">Link has also been saved to your "Settings" sheet.</p>
    </body>
    </html>
  `;
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Launch Dashboard');
}

/**
 * Fetch initial database contents and status statistics
 */
function getInitialData() {
  try {
    initializeSheetsIfNeeded();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = getSettingsSheet(ss);
    var membersSheet = ss.getSheetByName("Members");
    var paymentsSheet = ss.getSheetByName("Payments Received");
    var expensesSheet = ss.getSheetByName("Expenses");
    var accountsSheet = ss.getSheetByName("Accounts");
    var fdSheet = ss.getSheetByName("Fixed Deposits");
    var usersSheet = ss.getSheetByName("Users");
    var pettySheet = ss.getSheetByName("Petty Cash Advances");
    
    // Automatically check and add next financial year column if past March 31st
    checkAndAddNextFinancialYear(membersSheet);
    
    var settings = getSettingsMap(settingsSheet);
    
    // Dynamic URL check: only save if missing or not a valid URL
    try {
      if (!settings["Web App Link"] || settings["Web App Link"].indexOf("http") === -1) {
        var webAppUrl = ScriptApp.getService().getUrl();
        if (webAppUrl) {
          updateSettingValue("Web App Link", webAppUrl);
          settings["Web App Link"] = webAppUrl;
        }
      }
    } catch(e) {}
    
    var membersRange = membersSheet.getDataRange();
    var membersValues = membersRange.getValues();
    var membersHeaders = membersValues[0];
    
    var memberIdIdx = membersHeaders.indexOf("Member ID");
    var nameIdx = membersHeaders.indexOf("Name");
    var mobileIdx = membersHeaders.indexOf("Mobile Number");
    var famCountIdx = membersHeaders.indexOf("No of fam mem");
    var joinDateIdx = membersHeaders.indexOf("Join Date");
    var srNoIdx = membersHeaders.indexOf("Sr No");
    
    // Parse years dynamically: columns whose header matches yyyy-yy Status (e.g. 2022-23 Status)
    var financialYears = [];
    var yearColIndices = [];
    var yearReceiptIndices = [];
    
    membersHeaders.forEach(function(h, idx) {
      if (h.indexOf(" Status") > -1) {
        var yr = h.replace(" Status", "").trim();
        var startYear = parseInt(yr.substring(0, 4));
        if (startYear >= 2022) {
          financialYears.push(yr);
          yearColIndices.push(idx);
          yearReceiptIndices.push(membersHeaders.indexOf(yr + " Receipt"));
        }
      }
    });
    
    // Parse Members List
    var membersList = [];
    for (var i = 1; i < membersValues.length; i++) {
      if (!membersValues[i][nameIdx]) continue;
      
      var member = {};
      member["Sr No"] = membersValues[i][srNoIdx];
      member["Member ID"] = membersValues[i][memberIdIdx];
      member["Name"] = membersValues[i][nameIdx];
      member["Mobile Number"] = membersValues[i][mobileIdx];
      member["Family Members Count"] = membersValues[i][famCountIdx] || 1;
      
      var joinDateVal = membersValues[i][joinDateIdx];
      if (joinDateVal instanceof Date) {
        joinDateVal = formatDateString(joinDateVal);
      }
      member["Join Date"] = joinDateVal;
      
      // Map year statuses dynamically
      financialYears.forEach(function(yr, yIdx) {
        var statusColIdx = yearColIndices[yIdx];
        var receiptColIdx = yearReceiptIndices[yIdx];
        
        var rawStatus = membersValues[i][statusColIdx];
        var rawReceipt = receiptColIdx > -1 ? membersValues[i][receiptColIdx] : "";
        
        var isPaid = false;
        var isNA = false;
        
        if (rawStatus !== undefined && rawStatus !== null && rawStatus !== "") {
          var valStr = rawStatus.toString().trim().toUpperCase();
          if (valStr === "NA" || valStr === "N/A") {
            isNA = true;
          } else if (valStr === "TRUE" || rawStatus === true) {
            isPaid = true;
          }
        }
        
        member[yr + " Status"] = isNA ? "NA" : (isPaid ? "Paid" : "");
        member[yr + " Receipt"] = rawReceipt ? rawReceipt.toString().trim() : "";
      });
      
      membersList.push(member);
    }
    
    // Parse Payments
    var paymentsRange = paymentsSheet.getDataRange();
    var paymentsValues = paymentsRange.getValues();
    var paymentsHeaders = paymentsValues[0];
    var paymentsList = [];
    for (var i = 1; i < paymentsValues.length; i++) {
      var payment = {};
      for (var j = 0; j < paymentsHeaders.length; j++) {
        var val = paymentsValues[i][j];
        if (val instanceof Date) val = formatDateString(val);
        payment[paymentsHeaders[j]] = val;
      }
      paymentsList.push(payment);
    }
    
    // Parse Expenses
    var expensesRange = expensesSheet.getDataRange();
    var expensesValues = expensesRange.getValues();
    var expensesHeaders = expensesValues[0];
    var expensesList = [];
    for (var i = 1; i < expensesValues.length; i++) {
      var expense = {};
      for (var j = 0; j < expensesHeaders.length; j++) {
        var val = expensesValues[i][j];
        if (val instanceof Date) val = formatDateString(val);
        expense[expensesHeaders[j]] = val;
      }
      expensesList.push(expense);
    }
    
    // Parse Accounts
    var accountsList = [];
    if (accountsSheet) {
      var accValues = accountsSheet.getDataRange().getValues();
      for (var i = 1; i < accValues.length; i++) {
        if (accValues[i][0]) {
          accountsList.push({
            name: accValues[i][0],
            type: accValues[i][1],
            openingBalance: parseFloat(accValues[i][2] || 0),
            currentBalance: parseFloat(accValues[i][3] || 0)
          });
        }
      }
    }
    
    // Parse Fixed Deposits
    var fdList = [];
    if (fdSheet) {
      var fdValues = fdSheet.getDataRange().getValues();
      var fdHeaders = fdValues[0];
      for (var i = 1; i < fdValues.length; i++) {
        if (fdValues[i][1]) {
          var fd = {};
          for (var j = 0; j < fdHeaders.length; j++) {
            var val = fdValues[i][j];
            if (val instanceof Date) val = formatDateString(val);
            fd[fdHeaders[j]] = val;
          }
          fdList.push(fd);
        }
      }
    }

    // Parse User profiles
    var usersList = [];
    if (usersSheet) {
      var userValues = usersSheet.getDataRange().getValues();
      for (var i = 1; i < userValues.length; i++) {
        if (userValues[i][0]) {
          usersList.push({
            username: userValues[i][0],
            displayName: userValues[i][2] || userValues[i][0]
          });
        }
      }
    }

    // Parse Petty Cash Advances
    var pettyList = [];
    if (pettySheet) {
      var pettyValues = pettySheet.getDataRange().getValues();
      var pettyHeaders = pettyValues[0];
      for (var i = 1; i < pettyValues.length; i++) {
        if (pettyValues[i][1]) {
          var adv = {};
          for (var j = 0; j < pettyHeaders.length; j++) {
            var val = pettyValues[i][j];
            if (val instanceof Date) val = formatDateString(val);
            adv[pettyHeaders[j]] = val;
          }
          pettyList.push(adv);
        }
      }
    }
    
    // Calculate statistics
    var totalMembers = membersList.length;
    var totalCollected = 0;
    paymentsList.forEach(function(p) {
      if (p["Receipt Number"] !== "DELETED") {
        totalCollected += parseFloat(p["Total Amount"] || 0);
      }
    });
    
    var totalExpenses = 0;
    expensesList.forEach(function(e) {
      if (e["Voucher Number"] !== "DELETED") {
        totalExpenses += parseFloat(e["Amount"] || 0);
      }
    });
    
    var outstandingCount = 0;
    membersList.forEach(function(m) {
      var hasPending = false;
      financialYears.forEach(function(yr) {
        var status = m[yr + " Status"];
        if (status === undefined || status === null) status = "";
        var statusStr = status.toString().trim().toUpperCase();
        var isPaid = (statusStr === "TRUE" || statusStr === "PAID" || status === true);
        var isNA = (statusStr === "NA" || statusStr === "N/A");
        if (!isPaid && !isNA) {
          hasPending = true;
        }
      });
      if (hasPending) outstandingCount++;
    });
    
    return {
      success: true,
      settings: settings,
      members: membersList,
      payments: paymentsList,
      expenses: expensesList,
      accounts: accountsList,
      fixedDeposits: fdList,
      pettyCashAdvances: pettyList,
      users: usersList,
      years: financialYears,
      stats: {
        totalMembers: totalMembers,
        totalCollected: totalCollected,
        totalExpenses: totalExpenses,
        outstandingCount: outstandingCount
      }
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Save new Payment, record to ledger and generate receipt
 */
function savePayment(paymentData) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = getSettingsSheet(ss);
    var paymentsSheet = ss.getSheetByName("Payments Received");
    
    var settings = getSettingsMap(settingsSheet);
    var dateVal = paymentData.date ? new Date(paymentData.date) : new Date();
    var txYear = dateVal.getFullYear().toString();
    
    var prefix = settings["Receipt Prefix"] || "REC";
    var counter = parseInt(settings["Receipt Counter"] || "1000") + 1;
    updateSettingValue("Receipt Counter", counter);
    
    var receiptNo = paymentData.receiptOverride || (prefix + "-" + txYear + "-" + counter);
    
    var totalAmount = parseFloat(paymentData.donationAmount || 0);
    var breakdownArray = [];
    if (paymentData.yearlyFees && typeof paymentData.yearlyFees === 'object') {
      for (var yr in paymentData.yearlyFees) {
        var fee = parseFloat(paymentData.yearlyFees[yr]);
        totalAmount += fee;
        breakdownArray.push(yr + ": " + fee);
      }
    }
    var breakdownStr = breakdownArray.join("; ");
    
    // --- PDF & Drive Section (gracefully handles DriveApp access denied) ---
    var shortUrl = "";
    var pdfFileId = "";
    try {
      var folders = getFolderStructure();
      var yearFolder = getOrCreateFolder(folders.receipts, txYear);
      
      // Generate PDF Receipt
      var pdfName = receiptNo + "-" + sanitizeFilename(paymentData.name) + "-" + formatDateStringDdMmYy(dateVal);
      var pdfBlob = generatePDFBlob(receiptNo, paymentData, breakdownArray, settings, "Receipt");
      var pdfFile = yearFolder.createFile(pdfBlob).setName(pdfName + ".pdf");
      try {
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(sharingErr) {
        Logger.log("Could not set sharing on payment PDF: " + sharingErr.toString());
      }
      
      var longUrl = pdfFile.getUrl();
      shortUrl = shortenUrl(longUrl);
      pdfFileId = pdfFile.getId();
    } catch(driveErr) {
      // DriveApp access denied — payment is still recorded but PDF skipped
      // To fix: Open Script Editor > Run any function once to re-authorize Drive scope
      // OR add Drive scope to appsscript.json and redeploy
      shortUrl = "PDF_SKIPPED: " + driveErr.message;
      pdfFileId = "";
      Logger.log("DriveApp Error in savePayment: " + driveErr.toString());
    }
    
    // Write payment row
    var nextRow = paymentsSheet.getLastRow() + 1;
    var srNo = nextRow - 1;
    
    var headers = paymentsSheet.getRange(1, 1, 1, paymentsSheet.getLastColumn()).getValues()[0];
    var newRow = [];
    for (var col = 0; col < headers.length; col++) {
      newRow.push("");
    }
    
    newRow[0] = srNo;
    newRow[1] = receiptNo;
    newRow[2] = dateVal;
    newRow[3] = paymentData.name;
    newRow[4] = paymentData.memberId || "NON-MEMBER";
    newRow[5] = paymentData.mobileNumber || "";
    newRow[6] = paymentData.paymentAccount || "Cash";
    newRow[7] = paymentData.transactionId || "";
    newRow[8] = breakdownStr;
    newRow[9] = paymentData.donationAmount || 0;
    newRow[10] = totalAmount;
    newRow[11] = shortUrl;
    newRow[12] = pdfFileId;
    
    var prepByIdx = headers.indexOf("Prepared By");
    if (prepByIdx > -1) {
      newRow[prepByIdx] = paymentData.preparedBy || "system";
    }
    
    paymentsSheet.appendRow(newRow);
    
    // Post to ledger tab
    postToLedger(
      paymentData.paymentAccount || "Cash", 
      dateVal, 
      receiptNo, 
      "Received from " + paymentData.name + (paymentData.memberId && paymentData.memberId !== "NON-MEMBER" ? " (" + paymentData.memberId + ")" : "") + (breakdownStr ? " for dues: " + breakdownStr : "") + (paymentData.donationAmount > 0 ? " (Donation: " + paymentData.donationAmount + ")" : ""), 
      true, 
      totalAmount, 
      paymentData.preparedBy
    );
    
    if (paymentData.memberId && paymentData.memberId !== "NON-MEMBER") {
      updateMemberFeeStatusAndMobile(paymentData.memberId, paymentData.yearlyFees, "Paid", receiptNo, paymentData.mobileNumber);
    }
    
    // Warn if PDF was skipped due to Drive access issue
    if (shortUrl.indexOf("PDF_SKIPPED") === 0) {
      return { 
        success: true, 
        receiptNo: receiptNo, 
        shortUrl: "#", 
        warning: "Payment saved but PDF receipt could not be generated. Please re-authorize Drive access: open Script Editor > Run any function > Authorize when prompted, then redeploy." 
      };
    }
    
    return { success: true, receiptNo: receiptNo, shortUrl: shortUrl };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Save new Expense, record to ledger and generate voucher
 */
function saveExpense(expenseData) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = getSettingsSheet(ss);
    var expensesSheet = ss.getSheetByName("Expenses");
    
    var settings = getSettingsMap(settingsSheet);
    var dateVal = expenseData.date ? new Date(expenseData.date) : new Date();
    var txYear = dateVal.getFullYear().toString();
    
    var prefix = settings["Voucher Prefix"] || "EXP";
    var counter = parseInt(settings["Voucher Counter"] || "1000") + 1;
    updateSettingValue("Voucher Counter", counter);
    
    var voucherNo = expenseData.voucherOverride || (prefix + "-" + txYear + "-" + counter);
    
    // --- PDF & Drive Section (gracefully handles DriveApp access denied) ---
    var shortUrl = "";
    var pdfFileId = "";
    try {
      var folders = getFolderStructure();
      var yearFolder = getOrCreateFolder(folders.vouchers, txYear);
      
      // Generate PDF Voucher
      var pdfName = voucherNo + "-" + sanitizeFilename(expenseData.paidTo) + "-" + formatDateStringDdMmYy(dateVal);
      var pdfBlob = generatePDFBlob(voucherNo, expenseData, [], settings, "Payment Voucher");
      var pdfFile = yearFolder.createFile(pdfBlob).setName(pdfName + ".pdf");
      try {
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(sharingErr) {
        Logger.log("Could not set sharing on expense PDF: " + sharingErr.toString());
      }
      
      var longUrl = pdfFile.getUrl();
      shortUrl = shortenUrl(longUrl);
      pdfFileId = pdfFile.getId();
    } catch(driveErr) {
      shortUrl = "PDF_SKIPPED: " + driveErr.message;
      pdfFileId = "";
      Logger.log("DriveApp Error in saveExpense: " + driveErr.toString());
    }
    
    // Write expense row
    var nextRow = expensesSheet.getLastRow() + 1;
    var srNo = nextRow - 1;
    
    var headers = expensesSheet.getRange(1, 1, 1, expensesSheet.getLastColumn()).getValues()[0];
    var newRow = [];
    for (var col = 0; col < headers.length; col++) {
      newRow.push("");
    }
    
    newRow[0] = srNo;
    newRow[1] = voucherNo;
    newRow[2] = dateVal;
    newRow[3] = expenseData.paidTo;
    newRow[4] = expenseData.amount || 0;
    newRow[5] = expenseData.category || "General";
    newRow[6] = expenseData.narration || "";
    newRow[7] = expenseData.paymentAccount || "Cash";
    newRow[8] = expenseData.referenceId || "";
    newRow[9] = shortUrl;
    newRow[10] = pdfFileId;
    
    var prepByIdx = headers.indexOf("Prepared By");
    if (prepByIdx > -1) {
      newRow[prepByIdx] = expenseData.preparedBy || "system";
    }
    
    expensesSheet.appendRow(newRow);
    
    // Post to ledger tab
    postToLedger(
      expenseData.paymentAccount || "Cash", 
      dateVal, 
      voucherNo, 
      "Paid to " + expenseData.paidTo + " for " + expenseData.category + " (" + expenseData.narration + ")", 
      false, 
      expenseData.amount, 
      expenseData.preparedBy
    );
    
    if (shortUrl.indexOf("PDF_SKIPPED") === 0) {
      return { success: true, voucherNo: voucherNo, shortUrl: "#", warning: "Expense saved but PDF voucher could not be generated. Re-authorize Drive access in Script Editor." };
    }
    
    return { success: true, voucherNo: voucherNo, shortUrl: shortUrl };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Edit existing Payment and adjust ledgers
 */
function editPayment(receiptNo, paymentData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paymentsSheet = ss.getSheetByName("Payments Received");
    var settingsSheet = getSettingsSheet(ss);
    var settings = getSettingsMap(settingsSheet);
    
    var dataRange = paymentsSheet.getDataRange();
    var values = dataRange.getValues();
    var rowIndex = -1;
    var oldPaymentRow = null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === receiptNo) {
        rowIndex = i + 1;
        oldPaymentRow = values[i];
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Receipt not found: " + receiptNo);
    
    var oldMemberId = oldPaymentRow[4];
    var oldBreakdownStr = oldPaymentRow[8];
    var oldPdfId = oldPaymentRow[12];
    var oldAccount = oldPaymentRow[6];
    
    if (oldMemberId && oldMemberId !== "NON-MEMBER" && oldBreakdownStr) {
      var oldFees = parseBreakdownString(oldBreakdownStr);
      updateMemberFeeStatusAndMobile(oldMemberId, oldFees, "", "", "");
    }
    
    if (oldPdfId) {
      try {
        DriveApp.getFileById(oldPdfId).setTrashed(true);
      } catch(e) {}
    }
    
    var dateVal = paymentData.date ? new Date(paymentData.date) : new Date();
    var txYear = dateVal.getFullYear().toString();
    
    var totalAmount = parseFloat(paymentData.donationAmount || 0);
    var breakdownArray = [];
    if (paymentData.yearlyFees && typeof paymentData.yearlyFees === 'object') {
      for (var yr in paymentData.yearlyFees) {
        var fee = parseFloat(paymentData.yearlyFees[yr]);
        totalAmount += fee;
        breakdownArray.push(yr + ": " + fee);
      }
    }
    var breakdownStr = breakdownArray.join("; ");
    
    var folders = getFolderStructure();
    var yearFolder = getOrCreateFolder(folders.receipts, txYear);
    var pdfName = receiptNo + "-" + sanitizeFilename(paymentData.name) + "-" + formatDateStringDdMmYy(dateVal);
    var pdfBlob = generatePDFBlob(receiptNo, paymentData, breakdownArray, settings, "Receipt");
    var pdfFile = yearFolder.createFile(pdfBlob).setName(pdfName + ".pdf");
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingErr) {
      Logger.log("Could not set sharing on payment PDF: " + sharingErr.toString());
    }
    
    var longUrl = pdfFile.getUrl();
    var shortUrl = shortenUrl(longUrl);
    
    var headers = paymentsSheet.getRange(1, 1, 1, paymentsSheet.getLastColumn()).getValues()[0];
    
    paymentsSheet.getRange(rowIndex, 3, 1, 11).setValues([[
      dateVal,
      paymentData.name,
      paymentData.memberId || "NON-MEMBER",
      paymentData.mobileNumber || "",
      paymentData.paymentAccount || "Cash",
      paymentData.transactionId || "",
      breakdownStr,
      paymentData.donationAmount || 0,
      totalAmount,
      shortUrl,
      pdfFile.getId()
    ]]);
    
    var prepByIdx = headers.indexOf("Prepared By");
    if (prepByIdx > -1) {
      paymentsSheet.getRange(rowIndex, prepByIdx + 1).setValue(paymentData.preparedBy || "system");
    }
    
    var description = "Received from " + paymentData.name + (paymentData.memberId && paymentData.memberId !== "NON-MEMBER" ? " (" + paymentData.memberId + ")" : "") + (breakdownStr ? " for dues: " + breakdownStr : "") + (paymentData.donationAmount > 0 ? " (Donation: " + paymentData.donationAmount + ")" : "");
    if (oldAccount !== paymentData.paymentAccount) {
      removeLedgerEntry(oldAccount, receiptNo);
      postToLedger(paymentData.paymentAccount, dateVal, receiptNo, description, true, totalAmount, paymentData.preparedBy);
    } else {
      updateLedgerEntry(paymentData.paymentAccount, receiptNo, dateVal, description, true, totalAmount, paymentData.preparedBy);
    }
    
    if (paymentData.memberId && paymentData.memberId !== "NON-MEMBER") {
      updateMemberFeeStatusAndMobile(paymentData.memberId, paymentData.yearlyFees, "Paid", receiptNo, paymentData.mobileNumber);
    }
    
    return { success: true, receiptNo: receiptNo, shortUrl: shortUrl };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Edit existing Expense and adjust ledgers
 */
function editExpense(voucherNo, expenseData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var expensesSheet = ss.getSheetByName("Expenses");
    var settingsSheet = getSettingsSheet(ss);
    var settings = getSettingsMap(settingsSheet);
    
    var dataRange = expensesSheet.getDataRange();
    var values = dataRange.getValues();
    var rowIndex = -1;
    var oldPdfId = null;
    var oldAccount = null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === voucherNo) {
        rowIndex = i + 1;
        oldPdfId = values[i][10];
        oldAccount = values[i][7];
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Voucher not found: " + voucherNo);
    
    if (oldPdfId) {
      try {
        DriveApp.getFileById(oldPdfId).setTrashed(true);
      } catch(e) {}
    }
    
    var dateVal = expenseData.date ? new Date(expenseData.date) : new Date();
    var txYear = dateVal.getFullYear().toString();
    
    var folders = getFolderStructure();
    var yearFolder = getOrCreateFolder(folders.vouchers, txYear);
    var pdfName = voucherNo + "-" + sanitizeFilename(expenseData.paidTo) + "-" + formatDateStringDdMmYy(dateVal);
    var pdfBlob = generatePDFBlob(voucherNo, expenseData, [], settings, "Payment Voucher");
    var pdfFile = yearFolder.createFile(pdfBlob).setName(pdfName + ".pdf");
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingErr) {
      Logger.log("Could not set sharing on expense PDF: " + sharingErr.toString());
    }
    
    var longUrl = pdfFile.getUrl();
    var shortUrl = shortenUrl(longUrl);
    
    var headers = expensesSheet.getRange(1, 1, 1, expensesSheet.getLastColumn()).getValues()[0];
    
    expensesSheet.getRange(rowIndex, 3, 1, 9).setValues([[
      dateVal,
      expenseData.paidTo,
      expenseData.amount || 0,
      expenseData.category || "General",
      expenseData.narration || "",
      expenseData.paymentAccount || "Cash",
      expenseData.referenceId || "",
      shortUrl,
      pdfFile.getId()
    ]]);
    
    var prepByIdx = headers.indexOf("Prepared By");
    if (prepByIdx > -1) {
      expensesSheet.getRange(rowIndex, prepByIdx + 1).setValue(expenseData.preparedBy || "system");
    }
    
    var description = "Paid to " + expenseData.paidTo + " for " + expenseData.category + " (" + expenseData.narration + ")";
    if (oldAccount !== expenseData.paymentAccount) {
      removeLedgerEntry(oldAccount, voucherNo);
      postToLedger(expenseData.paymentAccount, dateVal, voucherNo, description, false, expenseData.amount, expenseData.preparedBy);
    } else {
      updateLedgerEntry(expenseData.paymentAccount, voucherNo, dateVal, description, false, expenseData.amount, expenseData.preparedBy);
    }
    
    return { success: true, voucherNo: voucherNo, shortUrl: shortUrl };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete (Void) Payment and remove from ledger
 */
function deletePayment(receiptNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paymentsSheet = ss.getSheetByName("Payments Received");
    
    var dataRange = paymentsSheet.getDataRange();
    var values = dataRange.getValues();
    var rowIndex = -1;
    var oldPayment = null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === receiptNo) {
        rowIndex = i + 1;
        oldPayment = values[i];
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Receipt not found");
    
    var memberId = oldPayment[4];
    var breakdownStr = oldPayment[8];
    var pdfId = oldPayment[12];
    var oldAccount = oldPayment[6];
    
    if (memberId && memberId !== "NON-MEMBER" && breakdownStr) {
      var oldFees = parseBreakdownString(breakdownStr);
      updateMemberFeeStatusAndMobile(memberId, oldFees, "", "", "");
    }
    
    if (pdfId) {
      try {
        DriveApp.getFileById(pdfId).setTrashed(true);
      } catch(e) {}
    }
    
    paymentsSheet.getRange(rowIndex, 4, 1, 10).setValues([[
      "[DELETED RECORD]",
      "",
      "",
      "",
      "",
      "",
      "",
      0,
      0,
      ""
    ]]);
    
    paymentsSheet.getRange(rowIndex, 2).setValue("DELETED");
    removeLedgerEntry(oldAccount, receiptNo);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete (Void) Expense Voucher and remove from ledger
 */
function deleteExpense(voucherNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var expensesSheet = ss.getSheetByName("Expenses");
    
    var dataRange = expensesSheet.getDataRange();
    var values = dataRange.getValues();
    var rowIndex = -1;
    var pdfId = null;
    var oldAccount = null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === voucherNo) {
        rowIndex = i + 1;
        pdfId = values[i][10];
        oldAccount = values[i][7];
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Voucher not found");
    
    if (pdfId) {
      try {
        DriveApp.getFileById(pdfId).setTrashed(true);
      } catch(e) {}
    }
    
    expensesSheet.getRange(rowIndex, 4, 1, 8).setValues([[
      "[DELETED RECORD]",
      0,
      "",
      "",
      "",
      "",
      "",
      ""
    ]]);
    
    expensesSheet.getRange(rowIndex, 2).setValue("DELETED");
    removeLedgerEntry(oldAccount, voucherNo);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Manage Member List: Add / Edit / Delete Member
 */
function saveMember(memberData, isNew) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var membersSheet = ss.getSheetByName("Members");
    
    var values = membersSheet.getDataRange().getValues();
    var headers = values[0];
    
    var nameIdx = headers.indexOf("Name");
    var mobileIdx = headers.indexOf("Mobile Number");
    var famCountIdx = headers.indexOf("No of fam mem");
    var memberIdIdx = headers.indexOf("Member ID");
    var joinDateIdx = headers.indexOf("Join Date");
    var srNoIdx = headers.indexOf("Sr No");
    
    if (nameIdx === -1 || famCountIdx === -1) {
      throw new Error("Missing critical columns: 'Name' or 'No of fam mem'");
    }
    
    if (isNew) {
      var prefix = "M-";
      var lastRow = membersSheet.getLastRow();
      
      var idNum = 1001;
      if (lastRow > 1) {
        var idValues = membersSheet.getRange(2, memberIdIdx + 1, lastRow - 1, 1).getValues();
        var maxId = 1000;
        for (var k = 0; k < idValues.length; k++) {
          var idVal = idValues[k][0].toString();
          var match = idVal.match(/\d+/);
          if (match) {
            var num = parseInt(match[0]);
            if (num > maxId) maxId = num;
          }
        }
        idNum = maxId + 1;
      }
      var memberId = prefix + idNum;
      
      var newRow = [];
      for (var col = 0; col < headers.length; col++) {
        newRow.push("");
      }
      
      if (srNoIdx > -1) newRow[srNoIdx] = lastRow;
      if (memberIdIdx > -1) newRow[memberIdIdx] = memberId;
      if (nameIdx > -1) newRow[nameIdx] = memberData.Name;
      if (mobileIdx > -1) newRow[mobileIdx] = memberData["Mobile Number"];
      if (famCountIdx > -1) newRow[famCountIdx] = memberData["Family Members Count"] || 1;
      if (joinDateIdx > -1) newRow[joinDateIdx] = new Date();
      
      headers.forEach(function(h, idx) {
        if (h.indexOf(" Status") > -1) {
          var yr = h.replace(" Status", "").trim();
          if (memberData.exemptions && memberData.exemptions.indexOf(yr) > -1) {
            newRow[idx] = "NA";
            var rIdx = headers.indexOf(yr + " Receipt");
            if (rIdx > -1) newRow[rIdx] = "NA";
          } else {
            newRow[idx] = false;
          }
        }
      });
      
      membersSheet.appendRow(newRow);
      return { success: true, memberId: memberId };
    } else {
      var rowIndex = -1;
      for (var i = 1; i < values.length; i++) {
        if (values[i][memberIdIdx] === memberData["Member ID"]) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex === -1) throw new Error("Member not found");
      
      if (nameIdx > -1) membersSheet.getRange(rowIndex, nameIdx + 1).setValue(memberData.Name);
      if (mobileIdx > -1) membersSheet.getRange(rowIndex, mobileIdx + 1).setValue(memberData["Mobile Number"]);
      if (famCountIdx > -1) membersSheet.getRange(rowIndex, famCountIdx + 1).setValue(memberData["Family Members Count"]);
      
      headers.forEach(function(h, idx) {
        if (h.indexOf(" Status") > -1) {
          var yr = h.replace(" Status", "").trim();
          var cell = membersSheet.getRange(rowIndex, idx + 1);
          var receiptCell = membersSheet.getRange(rowIndex, headers.indexOf(yr + " Receipt") + 1);
          
          if (memberData.exemptions && memberData.exemptions.indexOf(yr) > -1) {
            cell.setValue("NA");
            receiptCell.setValue("NA");
          } else {
            var currentVal = cell.getValue().toString().trim().toUpperCase();
            if (currentVal === "NA" || currentVal === "N/A") {
              cell.setValue(false);
              receiptCell.setValue("");
            }
          }
        }
      });
      
      return { success: true, memberId: memberData["Member ID"] };
    }
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function deleteMember(memberId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var membersSheet = ss.getSheetByName("Members");
    var dataRange = membersSheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[0];
    
    var memberIdIdx = headers.indexOf("Member ID");
    var rowIndex = -1;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][memberIdIdx] === memberId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Member not found");
    membersSheet.deleteRow(rowIndex);
    
    var lastRow = membersSheet.getLastRow();
    var srNoIdx = headers.indexOf("Sr No");
    if (lastRow > 1 && srNoIdx > -1) {
      for (var r = 2; r <= lastRow; r++) {
        membersSheet.getRange(r, srNoIdx + 1).setValue(r - 1);
      }
    }
    return { success: true };
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Settings updates from UI
 */
function updateSettings(settingsMap) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    for (var key in settingsMap) {
      updateSettingValue(key, settingsMap[key]);
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/* ========================================================
   Multi-Account Ledger Actions
   ======================================================== */

function addBankAccount(name, openingBalance) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accountsSheet = ss.getSheetByName("Accounts");
    var values = accountsSheet.getDataRange().getValues();
    
    var accountName = name.trim();
    if (!accountName || accountName.toUpperCase() === "USERS" || accountName.toUpperCase() === "SETTINGS" || accountName.toUpperCase() === "MEMBERS" || accountName.toUpperCase() === "FIXED DEPOSITS" || accountName.toUpperCase() === "PETTY CASH ADVANCES") {
      return { success: false, error: "Invalid account name!" };
    }
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString().trim().toLowerCase() === accountName.toLowerCase()) {
        return { success: false, error: "Account already exists!" };
      }
    }
    
    var opBal = parseFloat(openingBalance || 0);
    accountsSheet.appendRow([accountName, "Bank", opBal, opBal]);
    createLedgerSheet(accountName, opBal);
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getLedgerData(accountName) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(accountName);
    if (!sheet) throw new Error("Account ledger tab not found: " + accountName);
    
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var list = [];
    
    for (var i = 1; i < values.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var val = values[i][j];
        if (val instanceof Date) val = formatDateString(val);
        row[headers[j]] = val;
      }
      list.push(row);
    }
    
    return { success: true, ledger: list };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function saveTransfer(transferData) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var transfersSheet = ss.getSheetByName("Contra Transfers");
    if (!transfersSheet) {
      transfersSheet = ss.insertSheet("Contra Transfers");
      transfersSheet.appendRow(["Sr No", "Date", "From Account", "To Account", "Amount", "Reference/Notes", "Prepared By"]);
      transfersSheet.getRange("A1:G1").setFontWeight("bold");
    }
    
    var nextRow = transfersSheet.getLastRow() + 1;
    var srNo = nextRow - 1;
    var dateVal = transferData.date ? new Date(transferData.date) : new Date();
    var fromAcc = transferData.fromAccount;
    var toAcc = transferData.toAccount;
    var amount = parseFloat(transferData.amount || 0);
    var notes = transferData.notes || "";
    var preparedBy = transferData.preparedBy || "system";
    
    if (fromAcc === toAcc) {
      return { success: false, error: "Source and Destination accounts must be different!" };
    }
    
    transfersSheet.appendRow([
      srNo,
      dateVal,
      fromAcc,
      toAcc,
      amount,
      notes,
      preparedBy
    ]);
    
    var refNo = "TRF-" + srNo;
    postToLedger(toAcc, dateVal, refNo, "Transfer from " + fromAcc + " (" + notes + ")", true, amount, preparedBy);
    postToLedger(fromAcc, dateVal, refNo, "Transfer to " + toAcc + " (" + notes + ")", false, amount, preparedBy);
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/* ========================================================
   Petty Cash Advances Operations
   ======================================================== */

function saveAdvance(advanceData) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pettySheet = ss.getSheetByName("Petty Cash Advances");
    
    var dateVal = advanceData.date ? new Date(advanceData.date) : new Date();
    var nextRow = pettySheet.getLastRow() + 1;
    var srNo = nextRow - 1;
    var refNo = "ADV-" + srNo;
    
    pettySheet.appendRow([
      srNo,
      refNo,
      dateVal,
      advanceData.paidTo,
      parseFloat(advanceData.amountHandedOver || 0),
      advanceData.accountDebited,
      "Active",
      "", // Settled Date
      0,  // Total Spent
      0,  // Returned Amount
      "", // Voucher Number
      advanceData.notes || "",
      advanceData.preparedBy || "system"
    ]);
    
    // Post Debit/Credit entry in chosen ledger as Credit (Expense)
    postToLedger(
      advanceData.accountDebited,
      dateVal,
      refNo,
      "Petty Cash Advance handed to " + advanceData.paidTo + (advanceData.notes ? " (" + advanceData.notes + ")" : ""),
      false,
      parseFloat(advanceData.amountHandedOver || 0),
      advanceData.preparedBy
    );
    
    return { success: true, refNo: refNo };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function settleAdvance(settleData) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pettySheet = ss.getSheetByName("Petty Cash Advances");
    var values = pettySheet.getDataRange().getValues();
    
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === settleData.advanceRef) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Advance record not found: " + settleData.advanceRef);
    
    var amtHandedOver = parseFloat(values[rowIndex-1][4] || 0);
    var accountDebited = values[rowIndex-1][5];
    var paidTo = values[rowIndex-1][3];
    
    var dateVal = settleData.date ? new Date(settleData.date) : new Date();
    var totalSpent = parseFloat(settleData.totalSpent || 0);
    var returnedAmount = amtHandedOver - totalSpent;
    
    var settingsSheet = getSettingsSheet(ss);
    var settings = getSettingsMap(settingsSheet);
    var txYear = dateVal.getFullYear().toString();
    var prefix = settings["Voucher Prefix"] || "EXP";
    var counter = parseInt(settings["Voucher Counter"] || "1000") + 1;
    updateSettingValue("Voucher Counter", counter);
    var voucherNo = prefix + "-" + txYear + "-" + counter;
    
    // Save to Petty Advances tab
    pettySheet.getRange(rowIndex, 8, 1, 5).setValues([[
      dateVal,
      totalSpent,
      returnedAmount,
      voucherNo,
      settleData.notes || ""
    ]]);
    pettySheet.getRange(rowIndex, 7).setValue("Settled");
    
    // Post returned cash to Return account
    var refNoSettlement = "SET-" + values[rowIndex-1][0];
    if (returnedAmount > 0) {
      postToLedger(
        settleData.returnAccount || "Cash",
        dateVal,
        refNoSettlement,
        "Unspent advance returned by " + paidTo + " (Ref: " + settleData.advanceRef + ")",
        true,
        returnedAmount,
        settleData.preparedBy
      );
    } else if (returnedAmount < 0) {
      var extraSpent = Math.abs(returnedAmount);
      postToLedger(
        settleData.returnAccount || "Cash",
        dateVal,
        refNoSettlement,
        "Additional payout for settlement by " + paidTo + " (Ref: " + settleData.advanceRef + ")",
        false,
        extraSpent,
        settleData.preparedBy
      );
    }
    
    // Generate a single voucher inside Expenses sheet for total spent
    var expensesSheet = ss.getSheetByName("Expenses");
    var expNextRow = expensesSheet.getLastRow() + 1;
    var expSrNo = expNextRow - 1;
    var narrationStr = "Petty Cash Settlement by " + paidTo + " (Ref: " + settleData.advanceRef + "). Items: " + settleData.itemsSummary;
    
    var expenseData = {
      date: settleData.date,
      paidTo: paidTo,
      amount: totalSpent,
      category: "General",
      paymentAccount: accountDebited,
      referenceId: settleData.advanceRef,
      narration: narrationStr,
      preparedBy: settleData.preparedBy
    };
    
    var folders = getFolderStructure();
    var yearFolder = getOrCreateFolder(folders.vouchers, txYear);
    var pdfName = voucherNo + "-Settlement_" + sanitizeFilename(paidTo) + "-" + formatDateStringDdMmYy(dateVal);
    
    var detailsArray = [];
    if (settleData.itemsBreakdown && typeof settleData.itemsBreakdown === 'object') {
      settleData.itemsBreakdown.forEach(function(item) {
        detailsArray.push(item.desc + ": " + parseFloat(item.amount).toFixed(2));
      });
    }
    
    var pdfBlob = generatePDFBlob(voucherNo, expenseData, detailsArray, settings, "Payment Voucher");
    var pdfFile = yearFolder.createFile(pdfBlob).setName(pdfName + ".pdf");
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingErr) {
      Logger.log("Could not set sharing on settlement PDF: " + sharingErr.toString());
    }
    
    var longUrl = pdfFile.getUrl();
    var shortUrl = shortenUrl(longUrl);
    
    var expHeaders = expensesSheet.getRange(1, 1, 1, expensesSheet.getLastColumn()).getValues()[0];
    var expNewRow = [];
    for (var col = 0; col < expHeaders.length; col++) {
      expNewRow.push("");
    }
    
    expNewRow[0] = expSrNo;
    expNewRow[1] = voucherNo;
    expNewRow[2] = dateVal;
    expNewRow[3] = paidTo;
    expNewRow[4] = totalSpent;
    expNewRow[5] = "General";
    expNewRow[6] = narrationStr;
    expNewRow[7] = accountDebited;
    expNewRow[8] = settleData.advanceRef;
    expNewRow[9] = shortUrl;
    expNewRow[10] = pdfFile.getId();
    
    var prepIdx = expHeaders.indexOf("Prepared By");
    if (prepIdx > -1) {
      expNewRow[prepIdx] = settleData.preparedBy || "system";
    }
    
    expensesSheet.appendRow(expNewRow);
    
    return { success: true, voucherNo: voucherNo, shortUrl: shortUrl };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/* ========================================================
   Multi-User Credentials Verification
   ======================================================== */

function verifyLogin(username, password) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName("Users");
    var values = usersSheet.getDataRange().getValues();
    
    for (var i = 1; i < values.length; i++) {
      var u = values[i][0];
      var p = values[i][1];
      var d = values[i][2];
      
      if (u && u.toString().trim() === username.trim() && p && p.toString().trim() === password.trim()) {
        return { success: true, user: { username: u, displayName: d } };
      }
    }
    return { success: false, error: "Invalid username or password!" };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getUsers() {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName("Users");
    var values = usersSheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < values.length; i++) {
      if (values[i][0]) {
        list.push({
          username: values[i][0],
          password: values[i][1],
          displayName: values[i][2]
        });
      }
    }
    return { success: true, users: list };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function saveUsers(usersList) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName("Users");
    
    var limit = Math.min(usersList.length, 3);
    var lastRow = usersSheet.getLastRow();
    if (lastRow > 1) {
      usersSheet.deleteRows(2, lastRow - 1);
    }
    
    for (var i = 0; i < limit; i++) {
      var u = usersList[i];
      usersSheet.appendRow([u.username.trim(), u.password.trim(), u.displayName.trim()]);
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/* ========================================================
   Fixed Deposits certificate operations
   ======================================================== */

function saveFD(fdData) {
  try {
    initializeSheetsIfNeeded();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fdSheet = ss.getSheetByName("Fixed Deposits");
    
    var values = fdSheet.getDataRange().getValues();
    var isNew = (!fdData.srNo);
    
    var startDateVal = fdData.startDate ? new Date(fdData.startDate) : new Date();
    var maturityDateVal = fdData.maturityDate ? new Date(fdData.maturityDate) : new Date();
    
    if (isNew) {
      var nextRow = fdSheet.getLastRow() + 1;
      var srNo = nextRow - 1;
      fdSheet.appendRow([
        srNo,
        fdData.fdNumber,
        fdData.bankName,
        parseFloat(fdData.principalAmount || 0),
        parseFloat(fdData.interestRate || 0),
        startDateVal,
        maturityDateVal,
        parseFloat(fdData.maturityAmount || 0),
        fdData.undersigned,
        fdData.notes,
        fdData.status || "Active"
      ]);
    } else {
      var rowIndex = -1;
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] == fdData.srNo) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex === -1) throw new Error("FD Certificate record not found");
      
      fdSheet.getRange(rowIndex, 2, 1, 10).setValues([[
        fdData.fdNumber,
        fdData.bankName,
        parseFloat(fdData.principalAmount || 0),
        parseFloat(fdData.interestRate || 0),
        startDateVal,
        maturityDateVal,
        parseFloat(fdData.maturityAmount || 0),
        fdData.undersigned,
        fdData.notes,
        fdData.status || "Active"
      ]]);
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function deleteFD(srNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fdSheet = ss.getSheetByName("Fixed Deposits");
    if (!fdSheet) throw new Error("Fixed Deposits sheet not found");
    
    var values = fdSheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == srNo) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Record not found");
    fdSheet.deleteRow(rowIndex);
    
    var lastRow = fdSheet.getLastRow();
    if (lastRow > 1) {
      for (var r = 2; r <= lastRow; r++) {
        fdSheet.getRange(r, 1).setValue(r - 1);
      }
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/* ========================================================
   Internal Ledger Helper Posting Functions
   ======================================================== */

function createLedgerSheet(name, openingBalance) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(["Date", "Ref Number", "Particulars", "Income (Dr)", "Expense (Cr)", "Balance", "Prepared By"]);
    sheet.getRange("A1:G1").setFontWeight("bold");
    
    var today = new Date();
    sheet.appendRow([today, "-", "Opening Balance", parseFloat(openingBalance || 0), 0, parseFloat(openingBalance || 0), "system"]);
    
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 110);
    sheet.setColumnWidth(7, 100);
  }
}

function postToLedger(accountName, date, refNo, particulars, isIncome, amount, preparedBy) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(accountName);
  if (!sheet) {
    var opBal = 0;
    var accountsSheet = ss.getSheetByName("Accounts");
    if (accountsSheet) {
      var vals = accountsSheet.getDataRange().getValues();
      for (var i = 1; i < vals.length; i++) {
        if (vals[i][0] === accountName) {
          opBal = parseFloat(vals[i][2] || 0);
          break;
        }
      }
    }
    createLedgerSheet(accountName, opBal);
    sheet = ss.getSheetByName(accountName);
  }
  
  var dateVal = date ? new Date(date) : new Date();
  var income = isIncome ? parseFloat(amount || 0) : 0;
  var expense = isIncome ? 0 : parseFloat(amount || 0);
  
  sheet.appendRow([dateVal, refNo, particulars, income, expense, 0, preparedBy || ""]);
  recalculateLedger(accountName);
}

function recalculateLedger(accountName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(accountName);
  if (!sheet) return 0;
  
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (values.length <= 1) return 0;
  
  var accountsSheet = ss.getSheetByName("Accounts");
  var openingBal = 0;
  var rowIdxInAccounts = -1;
  
  if (accountsSheet) {
    var accValues = accountsSheet.getDataRange().getValues();
    for (var i = 1; i < accValues.length; i++) {
      if (accValues[i][0] === accountName) {
        openingBal = parseFloat(accValues[i][2] || 0);
        rowIdxInAccounts = i + 1;
        break;
      }
    }
  }
  
  var currentBalance = openingBal;
  var updates = [];
  
  values[1][5] = openingBal;
  updates.push([openingBal]);
  
  for (var r = 2; r < values.length; r++) {
    var income = parseFloat(values[r][3] || 0);
    var expense = parseFloat(values[r][4] || 0);
    currentBalance = currentBalance + income - expense;
    values[r][5] = currentBalance;
    updates.push([currentBalance]);
  }
  
  if (updates.length > 0) {
    sheet.getRange(2, 6, updates.length, 1).setValues(updates);
  }
  
  if (accountsSheet && rowIdxInAccounts > -1) {
    accountsSheet.getRange(rowIdxInAccounts, 4).setValue(currentBalance);
  }
  
  return currentBalance;
}

function removeLedgerEntry(accountName, refNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(accountName);
  if (!sheet) return;
  
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][1] === refNumber) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  recalculateLedger(accountName);
}

function updateLedgerEntry(accountName, refNumber, date, particulars, isIncome, amount, preparedBy) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(accountName);
  if (!sheet) return;
  
  var range = sheet.getDataRange();
  var values = range.getValues();
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (values[i][1] === refNumber) {
      rowIndex = i + 1;
      break;
    }
  }
  
  var dateVal = date ? new Date(date) : new Date();
  var income = isIncome ? parseFloat(amount || 0) : 0;
  var expense = isIncome ? 0 : parseFloat(amount || 0);
  
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 7).setValues([[
      dateVal,
      refNumber,
      particulars,
      income,
      expense,
      0,
      preparedBy || ""
    ]]);
  } else {
    sheet.appendRow([dateVal, refNumber, particulars, income, expense, 0, preparedBy || ""]);
  }
  recalculateLedger(accountName);
}

function getSettingsMap(sheet) {
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var key = values[i][0];
    var val = values[i][1];
    if (key) map[key.toString().trim()] = val;
  }
  return map;
}

function updateSettingValue(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSettingsSheet(ss);
  if (!sheet) {
    initializeSheetsIfNeeded();
    sheet = getSettingsSheet(ss) || ss.getSheetByName("Settings");
  }
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  var foundIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim() === key) {
      foundIndex = i + 1;
      break;
    }
  }
  
  if (foundIndex > -1) {
    sheet.getRange(foundIndex, 2).setValue(value);
  } else {
    sheet.appendRow([key, value]);
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    var newFolder = parentFolder.createFolder(folderName);
    try {
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingErr) {
      Logger.log("Could not set sharing on folder " + folderName + ": " + sharingErr.toString());
    }
    return newFolder;
  }
}

function getFolderStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = getSettingsSheet(ss);
  var settings = getSettingsMap(settingsSheet);
  
  var orgName = settings["Organization Name"] || "Our Community Association";
  
  var driveFolderId = settings["Drive Folder ID"];
  var receiptsFolderId = settings["Receipts Folder ID"];
  var vouchersFolderId = settings["Vouchers Folder ID"];
  
  var driveFolder, receiptsFolder, vouchersFolder;
  
  try {
    if (driveFolderId) driveFolder = DriveApp.getFolderById(driveFolderId);
  } catch(e) {}
  
  try {
    if (receiptsFolderId) receiptsFolder = DriveApp.getFolderById(receiptsFolderId);
  } catch(e) {}
  
  try {
    if (vouchersFolderId) vouchersFolder = DriveApp.getFolderById(vouchersFolderId);
  } catch(e) {}
  
  // Create folders if missing
  if (!driveFolder) {
    var sheetFile = DriveApp.getFileById(ss.getId());
    var parents = sheetFile.getParents();
    var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    
    driveFolder = parentFolder.createFolder(orgName + " Documents");
    try {
      driveFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingErr) {
      Logger.log("Could not set sharing on drive folder: " + sharingErr.toString());
    }
    updateSettingValue("Drive Folder ID", driveFolder.getId());
  }
  
  if (!receiptsFolder) {
    receiptsFolder = getOrCreateFolder(driveFolder, "Receipts");
    updateSettingValue("Receipts Folder ID", receiptsFolder.getId());
  }
  
  if (!vouchersFolder) {
    vouchersFolder = getOrCreateFolder(driveFolder, "Vouchers");
    updateSettingValue("Vouchers Folder ID", vouchersFolder.getId());
  }
  
  return {
    drive: driveFolder,
    receipts: receiptsFolder,
    vouchers: vouchersFolder
  };
}

function shortenUrl(longUrl) {
  if (!longUrl || longUrl.indexOf("http") !== 0) return longUrl;
  
  // Try TinyURL first
  try {
    var response = UrlFetchApp.fetch("https://tinyurl.com/api-create.php?url=" + encodeURIComponent(longUrl), {
      muteHttpExceptions: true,
      connectTimeout: 5000,
      readTimeout: 5000
    });
    if (response.getResponseCode() === 200) {
      var result = response.getContentText().trim();
      if (result && result.indexOf("http") === 0) {
        return result;
      }
    }
  } catch (e) {
    Logger.log("TinyURL shortening failed: " + e.toString());
  }
  
  // Fallback to is.gd
  try {
    var response = UrlFetchApp.fetch("https://is.gd/create.php?format=simple&url=" + encodeURIComponent(longUrl), {
      muteHttpExceptions: true,
      connectTimeout: 5000,
      readTimeout: 5000
    });
    if (response.getResponseCode() === 200) {
      var result = response.getContentText().trim();
      if (result && result.indexOf("http") === 0) {
        return result;
      }
    }
  } catch (e) {
    Logger.log("is.gd shortening failed: " + e.toString());
  }
  
  return longUrl;
}

/**
 * Generate PDF blob from HTML Template
 */
function generatePDFBlob(number, data, detailsArray, settings, typeTitle) {
  var orgName = settings["Organization Name"] || "Our Community Association";
  var dateStr = data.date ? formatDateStringDdMmYy(new Date(data.date)) : formatDateStringDdMmYy(new Date());
  
  var name = data.name || data.paidTo || "Non-Member";
  
  var donationAmount = parseFloat(data.donationAmount || 0);
  var totalAmount = donationAmount;
  
  var detailsRowsHtml = "";
  if (typeTitle === "Receipt") {
    detailsArray.forEach(function(item) {
      var parts = item.split(": ");
      var yr = parts[0];
      var amt = parseFloat(parts[1] || 0);
      totalAmount += amt;
      detailsRowsHtml += `<tr>
        <td style="padding: 10px; border: 1px solid #ddd;">Yearly Membership Fee (${yr})</td>
        <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">INR ${amt.toFixed(2)}</td>
      </tr>`;
    });
    
    if (donationAmount > 0) {
      detailsRowsHtml += `<tr>
        <td style="padding: 10px; border: 1px solid #ddd;">Voluntary Donation</td>
        <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">INR ${donationAmount.toFixed(2)}</td>
      </tr>`;
    }
  } else {
    // Expense Voucher Itemized details
    totalAmount = parseFloat(data.amount || 0);
    if (detailsArray && detailsArray.length > 0) {
      detailsArray.forEach(function(item) {
        var parts = item.split(": ");
        var desc = parts[0];
        var amt = parseFloat(parts[1] || 0);
        detailsRowsHtml += `<tr>
          <td style="padding: 10px; border: 1px solid #ddd;">${desc}</td>
          <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">INR ${amt.toFixed(2)}</td>
        </tr>`;
      });
    } else {
      detailsRowsHtml = `<tr>
        <td style="padding: 10px; border: 1px solid #ddd;">Category: <strong>${data.category || 'General'}</strong><br/>Narration: ${data.narration || ''}</td>
        <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">INR ${totalAmount.toFixed(2)}</td>
      </tr>`;
    }
  }
  
  var htmlContent = `
  <html>
  <head>
    <style>
      body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 20px; color: #1F2937; }
      .container { border: 2px solid #374151; border-radius: 12px; padding: 30px; max-width: 650px; margin: auto; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
      .brand-header { text-align: center; border-bottom: 2px solid #E5E7EB; padding-bottom: 20px; }
      .org-name { font-size: 26px; font-weight: 800; color: #1E3A8A; letter-spacing: 0.5px; margin: 0; }
      .subtitle { font-size: 13px; color: #6B7280; text-transform: uppercase; margin-top: 5px; margin-bottom: 0; letter-spacing: 1px; }
      .title-box { display: inline-block; margin-top: 15px; background: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 20px; padding: 4px 18px; }
      .title-text { font-size: 14px; font-weight: 700; color: #1D4ED8; text-transform: uppercase; margin: 0; }
      
      .info-grid { width: 100%; margin-top: 25px; margin-bottom: 10px; border-collapse: collapse; }
      .info-grid td { padding: 4px 0; vertical-align: top; font-size: 14px; }
      .label { color: #6B7280; font-weight: 500; width: 30%; }
      .val { color: #1F2937; font-weight: 600; }
      
      .table-container { margin-top: 25px; }
      .details-table { width: 100%; border-collapse: collapse; font-size: 14px; }
      .details-table th { background: #F3F4F6; border: 1px solid #D1D5DB; padding: 10px; text-align: left; font-weight: 700; color: #374151; }
      .details-table td { border: 1px solid #E5E7EB; padding: 12px 10px; }
      .total-row td { font-weight: 800; font-size: 15px; border-top: 2px solid #374151; background: #F9FAFB; }
      
      .meta-footer { width: 100%; margin-top: 30px; border-collapse: collapse; font-size: 13px; }
      .meta-footer td { padding: 4px 0; }
      
      .signatures-section { width: 100%; margin-top: 60px; border-collapse: collapse; }
      .signatures-section td { text-align: center; font-size: 13px; color: #4B5563; }
      .line { width: 160px; border-bottom: 1px solid #4B5563; margin: 0 auto 8px auto; }
      
      .system-tag { text-align: center; margin-top: 40px; font-size: 11px; color: #9CA3AF; border-top: 1px solid #E5E7EB; padding-top: 10px; }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="brand-header">
        <h1 class="org-name">${orgName}</h1>
        <p class="subtitle">Official Transaction Record</p>
        <div class="title-box">
          <p class="title-text">${typeTitle}</p>
        </div>
      </div>
      
      <table class="info-grid">
        <tr>
          <td class="label">Number:</td>
          <td class="val">${number}</td>
          <td class="label" style="text-align: right; width: 20%;">Date:</td>
          <td class="val" style="text-align: right; width: 20%;">${dateStr}</td>
        </tr>
        <tr>
          <td class="label">${typeTitle === 'Receipt' ? 'Received From:' : 'Paid To:'}</td>
          <td class="val" colspan="3">${name} ${data.memberId && data.memberId !== 'NON-MEMBER' ? '(' + data.memberId + ')' : ''}</td>
        </tr>
        ${data.mobileNumber ? '<tr><td class="label">Mobile Number:</td><td class="val" colspan="3">' + data.mobileNumber + '</td></tr>' : ''}
      </table>
      
      <div class="table-container">
        <table class="details-table">
          <thead>
            <tr>
              <th>Description</th>
              <th style="text-align: right; width: 150px;">Amount</th>
            </tr>
          </thead>
          <tbody>
            ${detailsRowsHtml}
            <tr class="total-row">
              <td style="text-align: right; padding-right: 15px;">Total Amount:</td>
              <td style="text-align: right;">INR ${totalAmount.toFixed(2)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <table class="meta-footer">
        <tr>
          <td style="color: #6B7280; width: 30%;">Payment Method:</td>
          <td style="font-weight: 600;">${data.paymentAccount || 'Cash'}</td>
        </tr>
        ${data.transactionId || data.referenceId ? '<tr><td style="color: #6B7280;">Transaction/Ref ID:</td><td style="font-weight: 600;">' + (data.transactionId || data.referenceId) + '</td></tr>' : ''}
        ${data.narration && detailsArray && detailsArray.length > 0 ? '<tr><td style="color: #6B7280; vertical-align:top;">Narration:</td><td style="font-weight: 500; font-size:12px;">' + data.narration + '</td></tr>' : ''}
      </table>
      
      <table class="signatures-section">
        <tr>
          <td>
            <div class="line"></div>
            Prepared By: <strong>${data.preparedBy || 'system'}</strong>
          </td>
          <td>
            <div class="line"></div>
            Receiver Signature / Approved By
          </td>
        </tr>
      </table>
      
      <div class="system-tag">
        Generated electronically. Document verified under authority of ${orgName}.
      </div>
    </div>
  </body>
  </html>
  `;
  
  var tempFile = DriveApp.createFile("temp_receipt.html", htmlContent, MimeType.HTML);
  var pdfBlob = tempFile.getAs(MimeType.PDF);
  tempFile.setTrashed(true);
  return pdfBlob;
}

/**
 * Handle dynamic sheet structure initialization.
 */
function initializeSheetsIfNeeded() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Settings Tab
  var settingsSheet = getSettingsSheet(ss);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.appendRow(["Key", "Value"]);
    settingsSheet.appendRow(["Organization Name", "Excel Community Club"]);
    settingsSheet.appendRow(["Default Annual Fee", "1000"]);
    settingsSheet.appendRow(["Receipt Prefix", "REC"]);
    settingsSheet.appendRow(["Receipt Counter", "1000"]);
    settingsSheet.appendRow(["Voucher Prefix", "EXP"]);
    settingsSheet.appendRow(["Voucher Counter", "1000"]);
    settingsSheet.appendRow(["Drive Folder ID", ""]);
    settingsSheet.appendRow(["Receipts Folder ID", ""]);
    settingsSheet.appendRow(["Vouchers Folder ID", ""]);
    
    var waTemplate = "*{ORG}*\n\n" +
                     "Dear *{NAME}*,\n" +
                     "We have received your payment:\n\n" +
                     "• Receipt No: *{RECEIPT_NO}*\n" +
                     "• Amount: *INR {AMOUNT}*\n" +
                     "• Paid for: *{NARRATION}*\n\n" +
                     "Download Receipt: {LINK}\n\n" +
                     "Thanks,";
    settingsSheet.appendRow(["WhatsApp Template", waTemplate]);
    settingsSheet.appendRow(["Hidden Years", ""]);
    settingsSheet.appendRow(["Web App Link", ""]);
    settingsSheet.getRange("A1:B1").setFontWeight("bold");
    settingsSheet.setColumnWidth(1, 180);
    settingsSheet.setColumnWidth(2, 250);
  } else {
    // Add Hidden Years row if missing
    var settings = getSettingsMap(settingsSheet);
    if (settings["Hidden Years"] === undefined) {
      settingsSheet.appendRow(["Hidden Years", ""]);
    }
    // Add Web App Link row if missing
    if (settings["Web App Link"] === undefined) {
      settingsSheet.appendRow(["Web App Link", ""]);
    }
  }
  
  // Smart Migration: Check if Sheet1 exists, and Members does NOT exist
  var sheet1 = ss.getSheetByName("Sheet1");
  var membersSheet = ss.getSheetByName("Members");
  
  if (sheet1 && !membersSheet) {
    migrateSheet1ToMembers(ss, sheet1);
    membersSheet = ss.getSheetByName("Members");
  } else if (!membersSheet) {
    membersSheet = ss.insertSheet("Members");
    membersSheet.appendRow([
      "Sr No", "Member ID", "Name", "Mobile Number", "No of fam mem", "Join Date", 
      "2022-23 Status", "2022-23 Receipt", "2023-24 Status", "2023-24 Receipt", 
      "2024-25 Status", "2024-25 Receipt", "2025-26 Status", "2025-26 Receipt", 
      "2026-27 Status", "2026-27 Receipt"
    ]);
    membersSheet.getRange("A1:P1").setFontWeight("bold");
  }
  
  // Payments Tab
  var paymentsSheet = ss.getSheetByName("Payments Received");
  if (!paymentsSheet) {
    paymentsSheet = ss.insertSheet("Payments Received");
    paymentsSheet.appendRow([
      "Sr No", "Receipt Number", "Date", "Name", "Member ID", "Mobile Number", 
      "Payment Mode", "Transaction ID", "Fees Breakdown", "Donation Amount", "Total Amount", "Receipt Link", "Original PDF ID", "Prepared By"
    ]);
    paymentsSheet.getRange("A1:N1").setFontWeight("bold");
  } else {
    var pHeaders = paymentsSheet.getRange(1, 1, 1, paymentsSheet.getLastColumn()).getValues()[0];
    if (pHeaders.indexOf("Prepared By") === -1) {
      paymentsSheet.getRange(1, paymentsSheet.getLastColumn() + 1).setValue("Prepared By").setFontWeight("bold");
    }
  }
  
  // Expenses Tab
  var expensesSheet = ss.getSheetByName("Expenses");
  if (!expensesSheet) {
    expensesSheet = ss.insertSheet("Expenses");
    expensesSheet.appendRow([
      "Sr No", "Voucher Number", "Date", "Paid To", "Amount", "Category", 
      "Narration", "Payment Mode", "Reference ID", "Voucher Link", "Original PDF ID", "Prepared By"
    ]);
    expensesSheet.getRange("A1:L1").setFontWeight("bold");
  } else {
    var eHeaders = expensesSheet.getRange(1, 1, 1, expensesSheet.getLastColumn()).getValues()[0];
    if (eHeaders.indexOf("Prepared By") === -1) {
      expensesSheet.getRange(1, expensesSheet.getLastColumn() + 1).setValue("Prepared By").setFontWeight("bold");
    }
  }

  // Users Tab (Max 3 Users)
  var usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["Username", "Password", "Display Name"]);
    usersSheet.appendRow(["admin", "admin123", "Administrator"]);
    usersSheet.getRange("A1:C1").setFontWeight("bold");
  }

  // Accounts Tab
  var accountsSheet = ss.getSheetByName("Accounts");
  if (!accountsSheet) {
    accountsSheet = ss.insertSheet("Accounts");
    accountsSheet.appendRow(["Account Name", "Account Type", "Opening Balance", "Current Balance"]);
    accountsSheet.appendRow(["Cash", "Cash", 0, 0]);
    accountsSheet.getRange("A1:D1").setFontWeight("bold");
    createLedgerSheet("Cash", 0);
  }

  // Fixed Deposits Tab
  var fdSheet = ss.getSheetByName("Fixed Deposits");
  if (!fdSheet) {
    fdSheet = ss.insertSheet("Fixed Deposits");
    fdSheet.appendRow(["Sr No", "FD Number", "Bank Name", "Principal Amount", "Interest Rate (%)", "Start Date", "Maturity Date", "Maturity Amount", "Undersigned", "Notes", "Status"]);
    fdSheet.getRange("A1:K1").setFontWeight("bold");
  }

  // Petty Cash Advances Tab
  var pettySheet = ss.getSheetByName("Petty Cash Advances");
  if (!pettySheet) {
    pettySheet = ss.insertSheet("Petty Cash Advances");
    pettySheet.appendRow(["Sr No", "Advance Ref", "Date Given", "Paid To", "Amount Handed Over", "Account Debited", "Status", "Settled Date", "Total Spent", "Returned Amount", "Voucher Number", "Notes", "Prepared By"]);
    pettySheet.getRange("A1:M1").setFontWeight("bold");
  }
}

/**
 * Migration Runner: Sheet1 -> Members
 */
function migrateSheet1ToMembers(ss, sheet1) {
  var values = sheet1.getDataRange().getValues();
  var headers = values[0];
  
  var nameIdx = headers.indexOf("Name");
  var famCountIdx = headers.indexOf("No of fam mem");
  var srNoIdx = headers.indexOf("Sr No");
  
  if (nameIdx === -1 || famCountIdx === -1) {
    throw new Error("Cannot migrate Sheet1: missing 'Name' or 'No of fam mem' headers.");
  }
  
  var financialYears = [];
  var yearColIndices = [];
  headers.forEach(function(h, idx) {
    var headerTrimmed = h.trim();
    if (/^\d{4}-\d{2}$/.test(headerTrimmed)) {
      var startYear = parseInt(headerTrimmed.substring(0, 4));
      if (startYear >= 2022) {
        financialYears.push(headerTrimmed);
        yearColIndices.push(idx);
      }
    }
  });
  
  var membersHeaders = ["Sr No", "Member ID", "Name", "Mobile Number", "No of fam mem", "Join Date"];
  financialYears.forEach(function(yr) {
    membersHeaders.push(yr + " Status");
    membersHeaders.push(yr + " Receipt");
  });
  
  var membersSheet = ss.insertSheet("Members");
  membersSheet.appendRow(membersHeaders);
  membersSheet.getRange(1, 1, 1, membersHeaders.length).setFontWeight("bold");
  
  var outputRows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][nameIdx]) continue;
    
    var newRow = [];
    newRow.push(values[i][srNoIdx] || i);
    newRow.push("M-" + (1000 + i));
    newRow.push(values[i][nameIdx]);
    newRow.push("");
    newRow.push(values[i][famCountIdx] || 1);
    newRow.push(new Date());
    
    financialYears.forEach(function(yr, yIdx) {
      var colIdx = yearColIndices[yIdx];
      var val = values[i][colIdx];
      
      var statusVal = false;
      var receiptVal = "";
      
      if (val !== undefined && val !== null && val !== "") {
        var valStr = val.toString().trim().toUpperCase();
        if (valStr === "NA" || valStr === "N/A") {
          statusVal = "NA";
          receiptVal = "NA";
        } else if (valStr === "TRUE" || val === true) {
          statusVal = true;
          receiptVal = "NA";
        }
      }
      
      newRow.push(statusVal);
      newRow.push(receiptVal);
    });
    
    outputRows.push(newRow);
  }
  
  if (outputRows.length > 0) {
    membersSheet.getRange(2, 1, outputRows.length, membersHeaders.length).setValues(outputRows);
  }
  
  membersSheet.setColumnWidth(1, 50);
  membersSheet.setColumnWidth(2, 90);
  membersSheet.setColumnWidth(3, 180);
  membersSheet.setColumnWidth(4, 110);
  membersSheet.setColumnWidth(5, 120);
  membersSheet.setColumnWidth(6, 100);
  
  SpreadsheetApp.flush();
  
  try {
    sheet1.setName("Sheet1_Backup");
  } catch(e) {}
}

/**
 * Automatically checks the current date. If past March 31st (April 1st or later),
 * appends new financial year status and receipt columns.
 */
function checkAndAddNextFinancialYear(sheet) {
  var today = new Date();
  var year = today.getFullYear();
  var month = today.getMonth(); // Jan=0, Mar=2, Apr=3
  
  var currentFYStart;
  if (month >= 3) {
    currentFYStart = year;
  } else {
    currentFYStart = year - 1;
  }
  
  var endShort = (currentFYStart + 1) % 100;
  var endShortStr = endShort < 10 ? "0" + endShort : endShort.toString();
  var fyStr = currentFYStart + "-" + endShortStr;
  
  var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headers = headersRange.getValues()[0];
  
  var statusHeader = fyStr + " Status";
  var receiptHeader = fyStr + " Receipt";
  
  if (headers.indexOf(statusHeader) === -1) {
    var lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue(statusHeader).setFontWeight("bold");
    sheet.getRange(1, lastCol + 2).setValue(receiptHeader).setFontWeight("bold");
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var statusRange = sheet.getRange(2, lastCol + 1, lastRow - 1, 1);
      statusRange.setValue(false);
    }
    SpreadsheetApp.flush();
  }
}

function checkAndAddMissingHeaders(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var updated = false;
  
  var required = ["Member ID", "Mobile Number", "Join Date"];
  required.forEach(function(h) {
    if (headers.indexOf(h) === -1) {
      var lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue(h);
      headers.push(h);
      updated = true;
    }
  });
  
  if (updated) {
    SpreadsheetApp.flush();
  }
}

function formatDateString(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDateStringDdMmYy(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yy");
}

function sanitizeFilename(name) {
  if (!name) return "Unknown";
  return name.toString().replace(/[^a-zA-Z0-9\s-_]/g, "").trim().replace(/\s+/g, "_");
}

function updateMemberFeeStatusAndMobile(memberId, yearlyFees, statusValue, receiptNo, mobileNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var membersSheet = ss.getSheetByName("Members");
  var values = membersSheet.getDataRange().getValues();
  var headers = values[0];
  
  var memberIdIdx = headers.indexOf("Member ID");
  var mobileIdx = headers.indexOf("Mobile Number");
  
  var memberRowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (values[i][memberIdIdx] === memberId) {
      memberRowIndex = i + 1;
      break;
    }
  }
  
  if (memberRowIndex === -1) return;
  
  if (mobileNumber && mobileIdx > -1) {
    membersSheet.getRange(memberRowIndex, mobileIdx + 1).setValue(mobileNumber);
  }
  
  for (var yr in yearlyFees) {
    var statusColName = yr + " Status";
    var receiptColName = yr + " Receipt";
    
    var statusColIndex = headers.indexOf(statusColName);
    var receiptColIndex = headers.indexOf(receiptColName);
    
    if (statusColIndex === -1) {
      var newLastCol = membersSheet.getLastColumn() + 1;
      membersSheet.getRange(1, newLastCol).setValue(statusColName);
      headers.push(statusColName);
      statusColIndex = headers.length - 1;
      
      membersSheet.getRange(1, newLastCol + 1).setValue(receiptColName);
      headers.push(receiptColName);
      receiptColIndex = headers.length - 1;
    }
    
    var cellVal = false;
    if (statusValue === "Paid") {
      cellVal = true;
    } else if (statusValue === "NA") {
      cellVal = "NA";
    } else {
      cellVal = false;
    }
    
    membersSheet.getRange(memberRowIndex, statusColIndex + 1).setValue(cellVal);
    membersSheet.getRange(memberRowIndex, receiptColIndex + 1).setValue(receiptNo || "");
  }
}

function parseBreakdownString(breakdownStr) {
  var feesObj = {};
  if (!breakdownStr) return feesObj;
  var items = breakdownStr.split("; ");
  items.forEach(function(item) {
    var parts = item.split(": ");
    if (parts.length === 2) {
      feesObj[parts[0]] = parts[1];
    }
  });
  return feesObj;
}

function getSettingsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = ["Settings", "sittings", "sittng", "Sittings", "Sittng"];
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  return null;
}

/**
 * Handle incoming POST requests for standalone API connection.
 * Intercepts calls from standalone HTML dashboard and routes them to the proper backend function.
 */
function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    var functionName = request.functionName;
    var args = request.arguments || [];
    
    // Security Whitelist: Only expose necessary dashboard functions
    var allowedFunctions = [
      "getInitialData", 
      "savePayment", 
      "saveExpense", 
      "editPayment", 
      "editExpense",
      "deletePayment", 
      "deleteExpense", 
      "saveMember", 
      "deleteMember", 
      "updateSettings",
      "addBankAccount", 
      "getLedgerData", 
      "saveTransfer", 
      "saveAdvance", 
      "settleAdvance",
      "verifyLogin", 
      "getUsers", 
      "saveUsers", 
      "saveFD", 
      "deleteFD"
    ];
    
    if (allowedFunctions.indexOf(functionName) === -1) {
      throw new Error("Unauthorized or invalid backend function call: " + functionName);
    }
    
    if (typeof this[functionName] !== 'function') {
      throw new Error("Function " + functionName + " is not defined in the backend script.");
    }
    
    var result = this[functionName].apply(this, args);
    return ContentService.createTextOutput(JSON.stringify(result || { success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
