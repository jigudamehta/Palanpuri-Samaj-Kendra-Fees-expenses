# Membership & Expense Management System

A dual-mode responsive dashboard for managing members, collecting annual fees (with automated PDF receipt generation and WhatsApp reminders), and tracking program expenditures.

This project is structured to run **both** inside Google Sheets (container-bound mode) and as a **standalone website** (hosted on GitHub Pages or run locally) connecting to Google Sheets via a Web App API.

---

## 📂 Project Structure

- `Index.html`: The HTML5/CSS3/JavaScript frontend dashboard. Fully responsive with dark/light themes, tables, charts, forms, and dialogs.
- `Code.gs`: The Google Apps Script backend. Resolves database reads/writes from Google Sheets, manages user authentication, handles ledger postings, and compiles PDF receipts/vouchers in Google Drive.
- `appscript.json`: Manifest configuration file for Apps Script permissions and scopes.
- `.gitignore`: Configures Git to ignore system and temporary Google Drive files.

---

## 🚀 Setup & Deployment Guide

### Phase 1: Deploy Google Apps Script Web App (API Backend)
To allow the standalone HTML dashboard to connect to your Google Sheet, you must deploy the Google Apps Script as a Web App API:

1. Open your **Google Sheets** spreadsheet.
2. Go to the menu: **Extensions** > **Apps Script**.
3. Copy the contents of the local `Code.gs` and paste them into the script editor, replacing the existing code.
4. Copy the contents of `appscript.json` and paste them into the manifest file in Apps Script (if manifest files are hidden, click the gear icon `Project Settings` and check *Show "appsscript.json" manifest file in editor*).
5. Click the **Save** (disk) icon.
6. Click the **Deploy** button (top right) > **New deployment**.
7. Select type: **Web app** (click the gear icon next to "Select type" if Web App is not listed).
8. Configure the deployment settings:
   - **Description**: `Sheets Management API`
   - **Execute as**: **Me (your-email@gmail.com)** *(Crucial: This runs database operations using your spreadsheet & Drive permissions)*.
   - **Who has access**: **Anyone** *(Required to accept API calls from the standalone HTML dashboard)*.
9. Click **Deploy**.
10. If prompted, click **Authorize access**, select your Google account, click *Advanced*, and click *Go to Untitled project (unsafe)* to grant the script access to your spreadsheet and Google Drive folder.
11. Once completed, copy the **Web App URL** from the success screen (it will look like: `https://script.google.com/macros/s/AKfycb.../exec`).

> [!NOTE]
> If you make changes to `Code.gs` in the future, you must select **Deploy > Manage deployments** and edit the existing deployment to target the **New Version**, otherwise your Web App URL will continue running the old code version.

---

### Phase 2: Open and Configure Standalone Dashboard
Now you can open `Index.html` directly in a browser:

1. Double-click the `Index.html` file on your computer to open it in a browser, or upload it to a web host.
2. The sign-in overlay will appear, but since it is running standalone for the first time, it won't connect to the backend yet.
3. Click the **⚙️ Configure API Connection** link at the bottom of the sign-in card.
4. Paste the **Google Apps Script Web App URL** you copied in Phase 1 into the input field.
5. Click **Connect to Sheet**.
6. The dashboard will connect to Google Sheets, fetch the list of authorized users, and populate the *Username* dropdown.
7. Select your username, type your password (configured in the "Users" sheet of the Google Sheet, default mock credentials are `admin` / `admin123`), and click **Login**!

> [!TIP]
> You can update or replace the Web App URL at any time inside the **Settings** panel of the dashboard.

---

### Phase 3: PWA Installation & Forced Mobile Layout (Mobile Setup)
This application operates as a Progressive Web App (PWA) and offers a force-layout toggle, perfect for setting up 3-4 parallel mobile counters for fees collection:

1. **PWA App Installation**:
   - **Android (Chrome)**: Tap the menu (three dots) > **Install app** or follow the prompt to add the Samaj Manager dashboard to your home screen as a standalone application.
   - **iOS/iPhone (Safari)**: Tap the **Share** button (box with up-arrow) > scroll down and tap **Add to Home Screen**.
2. **Forced Mobile/Desktop Layout Toggle**:
   - A layout mode toggle button (smartphone icon) is added next to the dark-mode button in the header.
   - **On Desktop/Laptops**: Click it to force the **Mobile Layout** instantly. The interface collapses into a single column with scrollable tables and a drawer-style sidebar, allowing volunteers on laptops to work in the simplified mobile viewport.
   - **On Mobiles**: Tap it to force the **Desktop Layout** representation if you need to zoom in and see widescreen data views.

---

## 💻 Uploading to GitHub

Follow these steps to host your code on GitHub and enable GitHub Pages so you can access the dashboard from any device:

### 1. Initialize Git Repo
Open your terminal (PowerShell or Git Bash) inside the project folder:
```bash
git init
```

### 2. Commit Files
Add all files and make an initial commit:
```bash
git add .
git commit -m "Initial commit: Standalone dashboard & Sheets API connection"
```

### 3. Push to GitHub
Create a new repository on [GitHub](https://github.com/new) (keep it public if you want free GitHub Pages hosting). Then run the following (replacing with your repository details):
```bash
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

### 4. Enable GitHub Pages
1. Go to your repository on GitHub.
2. Navigate to **Settings** > **Pages** (under the Code and automation section).
3. Under *Build and deployment*, set **Source** to `Deploy from a branch`.
4. Under *Branch*, select `main` and `/ (root)` and click **Save**.
5. Wait 1-2 minutes. GitHub will generate a link for you, such as:
   `https://YOUR_USERNAME.github.io/YOUR_REPO_NAME/`
6. Open that link in your browser to access your premium standalone management dashboard anywhere!
