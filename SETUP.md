# CG Accounts Add-on Setup Guide

## Step 1: Create Standalone Project

Run this in your terminal (in the accounts-script folder):

```bash
npx @google/clasp create --type standalone --title "CG Accounts Add-on"
```

After it completes, you'll see a script URL. Note the **scriptId** from the output.

## Step 2: Update .clasp.json

After Step 1, a `.clasp.json` file will be created. Edit it to add rootDir:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "src"
}
```

## Step 3: Push Code

```bash
npx @google/clasp push
```

Type `y` if asked about overwriting.

## Step 4: Test Add-on

1. Run: `npx @google/clasp open`
2. In Apps Script editor: **Deploy → Test deployments**
3. Click **Install**
4. Open your Google Sheet
5. Go to **Extensions → CG Accounts Add-on**

## Step 5: Publish to Workspace

1. In Apps Script editor: **Deploy → New deployment**
2. Select type: **Add-on**
3. Fill description: "CG Accounts - Financial data aggregation"
4. Select: **Workspace internal only**
5. Click **Deploy**

Done! The add-on will be available to your classicgroup.asia domain.
