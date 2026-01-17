# CG Accounts Add-on - Deployment Guide

This guide walks you through deploying the CG Accounts add-on to your Google Workspace organization (classicgroup.asia).

## Prerequisites

- Node.js installed (v18+)
- npm installed
- Google account with access to classicgroup.asia Workspace
- Admin access to Google Workspace (for domain-wide installation)

## Overview

Deploying a Google Workspace Add-on requires:
1. Setting up a Google Cloud Platform (GCP) project
2. Configuring OAuth consent screen
3. Linking GCP project to Apps Script
4. Creating the add-on deployment
5. Installing domain-wide via Admin Console

---

## Step 1: Initial Setup (Local)

### 1.1 Clone and Install Dependencies

```bash
cd ~/Desktop/development/accounts-script
npm install
```

### 1.2 Login to Clasp

```bash
npx @google/clasp login
```

This opens a browser for Google OAuth. Login with your classicgroup.asia account.

### 1.3 Push Code to Apps Script

```bash
npm run push:force
```

### 1.4 Open Apps Script Editor

```bash
npx @google/clasp open
```

---

## Step 2: Create a GCP Project

The add-on requires a user-managed GCP project (not the default Apps Script-managed one).

### 2.1 Go to Google Cloud Console

Open [console.cloud.google.com](https://console.cloud.google.com)

### 2.2 Create New Project

1. Click the **project dropdown** at the top (next to "Google Cloud")
2. Click **New Project**
3. Enter:
   - **Project name**: `CG Accounts Add-on`
   - **Organization**: Select `classicgroup.asia`
   - **Location**: Choose your organization folder (or leave default)
4. Click **Create**
5. Wait for creation, then **select the new project**

### 2.3 Copy the Project Number

1. Go to **Dashboard** (or click the hamburger menu → Dashboard)
2. Find **Project number** (a number like `123456789012`)
3. **Copy this number** - you'll need it later

---

## Step 3: Configure OAuth Consent Screen

This is required for all add-ons.

### 3.1 Navigate to OAuth Consent Screen

1. In GCP Console, go to **APIs & Services** → **OAuth consent screen**
   - Or search "OAuth consent screen" in the search bar

### 3.2 Select User Type

1. Select **Internal** (for organization-only access)
2. Click **Create**

### 3.3 Fill in App Information

**App Information:**
- **App name**: `CG Accounts`
- **User support email**: Your email address
- **App logo**: (Optional) Upload a logo

**App Domain:** (Optional, can leave blank for internal apps)

**Developer contact information:**
- **Email addresses**: Your email address

Click **Save and Continue**

### 3.4 Configure Scopes

1. Click **Add or Remove Scopes**
2. In the filter/search, find and select these scopes:

| Scope | Description |
|-------|-------------|
| `https://www.googleapis.com/auth/spreadsheets` | Full access to spreadsheets |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | Access current spreadsheet |
| `https://www.googleapis.com/auth/script.container.ui` | Display UI elements |
| `https://www.googleapis.com/auth/gmail.send` | Send emails (for error notifications) |
| `https://www.googleapis.com/auth/drive.file` | Access Drive files created by app |

3. Click **Update**
4. Click **Save and Continue**

### 3.5 Summary

1. Review the summary
2. Click **Back to Dashboard**

---

## Step 4: Link GCP Project to Apps Script

### 4.1 Open Apps Script Project Settings

1. Go to your Apps Script editor: `npx @google/clasp open`
2. Click the **gear icon** ⚙️ (Project Settings) in the left sidebar

### 4.2 Change GCP Project

1. Scroll to **Google Cloud Platform (GCP) Project**
2. You'll see "GCP: Default"
3. Click **Change project**
4. Paste your **GCP Project Number** (from Step 2.3)
5. Click **Set project**

You should see a confirmation that the project is now linked.

---

## Step 5: Test the Add-on (Optional but Recommended)

Before deploying to everyone, test it yourself:

### 5.1 Create Test Deployment

1. In Apps Script editor, click **Deploy** → **Test deployments**
2. Click **Install**
3. This installs the add-on for your account only

### 5.2 Test Functionality

1. Open any Google Sheet
2. Go to **Extensions** → **CG Accounts**
3. Click **Open Dashboard**
4. Test Initialize, Test Connection, etc.

### 5.3 Uninstall Test (Optional)

1. Go back to **Deploy** → **Test deployments**
2. Click **Uninstall**

---

## Step 6: Create Production Deployment

### 6.1 New Deployment

1. In Apps Script editor, click **Deploy** → **New deployment**
2. Click the **gear icon** ⚙️ next to "Select type"
3. Select **Add-on**

### 6.2 Configure Deployment

- **Description**: `CG Accounts v1.0 - Financial Data Aggregation`

### 6.3 Deploy

1. Click **Deploy**
2. You'll see a success message with a **Deployment ID**
3. **Copy the Deployment ID** (looks like `AKfycbx...`)

Save this Deployment ID - you need it for the Admin Console installation.

---

## Step 7: Install Domain-Wide (Requires Workspace Admin)

### 7.1 Go to Admin Console

Open [admin.google.com](https://admin.google.com)

### 7.2 Navigate to Marketplace Apps

1. Click **Apps** in the left sidebar
2. Click **Google Workspace Marketplace apps**

### 7.3 Add Internal App

1. Click **Add app** (or the + button)
2. Select **Add internal app** (or "Add app to domain install allowlist")
3. Paste your **Deployment ID** from Step 6.3
4. Click **Continue** or **Add**

### 7.4 Configure Installation Scope

Choose who can use the add-on:
- **All users in organization** (recommended for this add-on)
- OR specific organizational units/groups

### 7.5 Install

1. Click **Install** or **Finish**
2. Wait for installation to complete

---

## Step 8: Verify Installation

### 8.1 Test with Your Account

1. Open any Google Sheet
2. Go to **Extensions** menu
3. You should see **CG Accounts** listed
4. Click **CG Accounts** → **Open Dashboard**

### 8.2 Test with Another User

Have another user in your organization:
1. Open any Google Sheet
2. Check that **CG Accounts** appears in Extensions menu

---

## Post-Deployment: First-Time Sheet Setup

For each ledger sheet that will use the add-on:

### 1. Initialize the Sheet

1. Open the target Google Sheet
2. Go to **Extensions** → **CG Accounts** → **Open Dashboard**
3. Click **Initialize Sheet**
4. This creates: CONFIG, Ledger Master, and RUN_LOG tabs

### 2. Configure the Sheet

Edit the **CONFIG** tab with your settings:

| Setting | Description | Example |
|---------|-------------|---------|
| ORG_CODE | Organization code | CM, KM, DEV, CPI, SM |
| ORG_NAME | Organization full name | Congzhou Machinery |
| FINANCIAL_YEAR | Current FY | 2025-26 |
| PURCHASE_SHEET_ID | Source sheet ID for purchases | 1abc...xyz |
| PURCHASE_SHEET_NAME | Tab name in source sheet | Purchase Register |
| SALES_SHEET_ID | Source sheet ID for sales | 1def...uvw |
| SALES_SHEET_NAME | Tab name in source sheet | Sales Register |
| BANK_SHEET_ID | Source sheet ID for bank | 1ghi...rst |
| BANK_SHEET_NAME | Tab name in source sheet | Bank Statement |
| ERROR_EMAIL | Email for error notifications | admin@classicgroup.asia |

### 3. Configure Column Mappings

In CONFIG tab, set the column names that match your source sheets:

**For Purchase Register:**
- PUR_DATE_COL, PUR_PARTY_ID_COL, PUR_PARTY_NAME_COL, PUR_AMOUNT_COL, etc.

**For Sales Register:**
- SAL_DATE_COL, SAL_PARTY_ID_COL, SAL_PARTY_NAME_COL, SAL_AMOUNT_COL, etc.

**For Bank Statement:**
- BANK_DATE_COL, BANK_PARTY_ID_COL, BANK_DEBIT_COL, BANK_CREDIT_COL, etc.

### 4. Test and Refresh

1. Click **Test Connection** to verify source sheet access
2. Click **Refresh Data** to fetch data and generate ledgers

---

## Updating the Add-on

When you make code changes:

### 1. Push Updated Code

```bash
npm run push:force
```

### 2. Create New Deployment Version

1. Open Apps Script: `npx @google/clasp open`
2. Click **Deploy** → **Manage deployments**
3. Click the **pencil icon** next to your deployment
4. Update the **Version** dropdown to a new version
5. Update the **Description** if needed
6. Click **Deploy**

The update will automatically roll out to all users.

---

## Troubleshooting

### "Apps Script-managed GCP project" Error

You need to link a user-managed GCP project. Follow Steps 2-4 above.

### "OAuth consent screen not configured" Error

Complete Step 3 (Configure OAuth Consent Screen) in GCP Console.

### Add-on Not Appearing in Extensions Menu

1. Make sure domain-wide installation completed (Step 7)
2. Try refreshing the Google Sheet
3. Check Admin Console to verify installation status

### "Permission denied" Errors

1. Verify OAuth scopes in GCP match those in appsscript.json
2. Re-authorize the add-on by removing and re-adding it

### Test Deployment Works but Production Doesn't

1. Create a new deployment version
2. Update the domain-wide installation with new deployment ID

---

## Quick Reference

### Useful Commands

```bash
# Push code to Apps Script
npm run push:force

# Open Apps Script editor
npx @google/clasp open

# View logs
npx @google/clasp logs

# Pull code from Apps Script (if edited online)
npx @google/clasp pull
```

### Important URLs

- **Apps Script**: https://script.google.com
- **GCP Console**: https://console.cloud.google.com
- **Admin Console**: https://admin.google.com
- **Script ID**: `1bCLTEW8-eb50xdSnqMGuVQ0rw6ncLZEuBRIJ6akFUuPGKHS1earfKG4D`

### Key Files

| File | Purpose |
|------|---------|
| `src/appsscript.json` | Add-on manifest with scopes and triggers |
| `src/main.js` | Entry points, menu, card UI |
| `src/init.js` | Sheet initialization and config |
| `.clasp.json` | Clasp configuration with Script ID |

---

## Support

For issues or questions:
1. Check the RUN_LOG tab for error details
2. View Apps Script logs: `npx @google/clasp logs`
3. Contact: [Your support email]
