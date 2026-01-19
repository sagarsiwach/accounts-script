# CG Accounts Script - Project Context

## Project Overview
Google Apps Script-based financial data aggregation system for Classic Group entities. Pulls from Purchase, Sales, and Bank source sheets and generates individual party ledgers.

## Tech Stack
- Runtime: Google Apps Script (V8)
- Local Dev: clasp + Node.js
- UI: HTML Service + Material Design Lite
- Version Control: Git + GitHub

## Key Architecture
- **Per-org design**: Each org's Ledger sheet is self-contained with CONFIG, Ledger Master, RUN_LOG tabs
- **Auto-detection**: Header rows are auto-detected by scanning for known columns (L/F, SUPPLIER NAME, etc.)
- **Bank multi-tab**: Scans all tabs in bank workbook, looking for schema in row 5
- **Contact-based company info**: Uses COMPANY_CONTACT_ID from CONFIG to pull company details from ALL CONTACTS sheet

## Tab Ordering
- Ledger Master (first)
- Individual party ledgers (CG-SUP-0001, CG-CUS-0001, etc.)
- CONFIG (second to last)
- RUN_LOG (last)

## Ledger Format Preferences
- Hide gridlines
- Column G hidden (stores IDs)
- Delete columns after G
- Column D width: 140px
- F1/G1: Light gray text (#999999)
- Row 2 (Company Name): UPPERCASE, BOLD
- Row 7 (Party Name): UPPERCASE, BOLD
- Currency format: Indian Rupees (₹)
- Font: Roboto Condensed throughout

## Contact Schema (ALL CONTACTS sheet)
- SL: Contact ID (CG-SUP-0001, CG-CUS-0001, CG-MAS-0001, etc.)
- CONTACT TYPE: SUPPLIER, CUSTOMER, MASTER, DEALER, CONTRACTOR, RENTAL
- COMPANY NAME, ADDRESS LINE 1, ADDRESS LINE 2, DISTRICT, STATE, PIN CODE
- GST, MOBILE NO, EMAIL ID

## Deployment Workflow

### Step 1: Push code and create version
```bash
clasp push && clasp version "Description of changes"
```
This outputs a version number (e.g., "Created version 8").

### Step 2: Update Google Cloud Marketplace SDK (REQUIRED)
After each `clasp version`, you MUST update the Marketplace SDK:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Select project: **cg-accounts** (or the linked GCP project)
3. Navigate to: **APIs & Services → Marketplace SDK → App Configuration**
4. Under **Google Workspace Add-on Deployment**:
   - Update **Version** to the new version number (e.g., 8)
5. Click **Save**

⚠️ **Without this step, users will still see the old version of the add-on!**

### Step 3: Git commit and push
```bash
git add -A && git commit -m "feat: Description" && git push origin master
```

### Quick Reference
| Action | Command/Location |
|--------|------------------|
| Push to Apps Script | `clasp push` |
| Create version | `clasp version "message"` |
| Update Marketplace | GCP Console → Marketplace SDK → App Configuration → Version |
| Current Script ID | `1unzAHnfpZuJM4sXnZW4JC8YLiD5MQDXm3rPU7xxB1O_hCamlEBnCMkJ_` |

## Organizations
| Code | Name |
|------|------|
| CM | Congzhou Machinery (Unipack) |
| KM | Kabira Mobility |
| DEV | DeltaEV |
| CPI | Classic Packaging Industry |
| SM | Satlok Machinery |
