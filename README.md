# CG Accounts Add-on

**Version:** 0.2.0
**Author:** Sagar Siwach
**Domain:** classicgroup.asia
**Last Updated:** 2026-01-16

---

## Overview

CG Accounts is a Google Sheets Add-on for aggregating financial data across Classic Group organizations. It pulls data from Purchase Registers, Sales Registers, and Bank Statements, then generates consolidated Ledger Masters and individual party ledgers.

### Key Features

- **Multi-org support**: Works on any sheet - one add-on for all organizations
- **Auto-initialization**: Creates CONFIG, Ledger Master, and RUN_LOG tabs automatically
- **Data aggregation**: Fetches from Purchase, Sales, Bank source sheets
- **Party ledgers**: Generates individual ledger sheets for each supplier/customer
- **Ledger Master index**: Hyperlinked list of all parties with balances
- **Logging**: Full execution history with error tracking
- **Email notifications**: Alerts on failures

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                    CG ACCOUNTS ADD-ON                               │
│              (Standalone - works on any sheet)                      │
└─────────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐     ┌───────────────┐     ┌───────────────┐
│ CM Ledger     │     │ KM Ledger     │     │ DEV Ledger    │
│ (any sheet)   │     │ (any sheet)   │     │ (any sheet)   │
├───────────────┤     ├───────────────┤     ├───────────────┤
│ CONFIG        │     │ CONFIG        │     │ CONFIG        │
│ Ledger Master │     │ Ledger Master │     │ Ledger Master │
│ RUN_LOG       │     │ RUN_LOG       │     │ RUN_LOG       │
│ CG-SUP-0001   │     │ CG-SUP-0001   │     │ CG-SUP-0001   │
│ CG-SUP-0002   │     │ ...           │     │ ...           │
│ ...           │     │               │     │               │
└───────────────┘     └───────────────┘     └───────────────┘
```

### Per-Sheet Structure

Each org's ledger sheet is self-contained:

| Tab | Purpose |
|-----|---------|
| `CONFIG` | All settings - org code, source sheet IDs, column mappings |
| `Ledger Master` | Index of all parties with hyperlinks and balances |
| `RUN_LOG` | Execution history with timestamps and status |
| `CG-SUP-XXXX` | Individual supplier ledger |
| `CG-CUST-XXXX` | Individual customer ledger |

---

## File Structure

```
accounts-script/
├── .clasp.json           # Clasp configuration
├── .claspignore          # Files to exclude from push
├── .gitignore            # Git exclusions
├── package.json          # NPM dependencies (Jest, clasp)
├── README.md             # This file
├── SETUP.md              # Quick setup guide
│
├── docs/
│   ├── roadmap.md        # Project roadmap and phases
│   └── sessions/
│       └── session-001.md # Development session log
│
├── src/                  # Google Apps Script source
│   ├── appsscript.json   # Apps Script manifest (add-on config)
│   ├── main.js           # Entry points, menu, triggers
│   ├── init.js           # Initialization - tab creation
│   ├── config.js         # Configuration reading (legacy)
│   ├── logger.js         # Run logging and email notifications
│   ├── utils.js          # Helper functions
│   │
│   ├── fetchers/
│   │   └── index.js      # Data fetchers (Purchase, Sales, Bank)
│   │
│   ├── ledgers/
│   │   ├── index.js      # Ledger generation (legacy)
│   │   └── party-ledger.js # Party ledger sheet creation
│   │
│   └── ui/
│       ├── Sidebar.html  # Main dashboard sidebar
│       ├── Logs.html     # Run logs viewer
│       ├── Settings.html # Settings dialog
│       └── ui-helpers.js # Server-side UI functions
│
└── tests/
    └── unit/
        ├── utils.test.js
        ├── fetchers.test.js
        └── ledgers.test.js
```

---

## Modules

### main.js

Entry point for the add-on. Contains:

| Function | Purpose |
|----------|---------|
| `onOpen(e)` | Creates CG Accounts menu |
| `onHomepage(e)` | Add-on card UI for sidebar |
| `createHomepageCard()` | Builds the add-on card |
| `refreshData()` | Main refresh - fetches all data and generates ledgers |
| `fetchSourceData(config, sourceType)` | Fetches from a single source |
| `transformRow(row, headers, mapping, sourceType)` | Transforms source row to standard format |
| `generateAllLedgers(ss, config, transactions)` | Creates all party ledger sheets |

### init.js

Initialization and configuration:

| Function | Purpose |
|----------|---------|
| `initialize()` | Creates all required tabs (CONFIG, Ledger Master, RUN_LOG) |
| `createConfigTab(ss)` | Creates CONFIG tab with all settings |
| `createLedgerMasterTab(ss)` | Creates Ledger Master index tab |
| `createRunLogTab(ss)` | Creates RUN_LOG tab |
| `readConfig()` | Reads configuration from CONFIG tab |
| `validateConfig()` | Validates configuration completeness |
| `testConnection()` | Tests connection to all source sheets |
| `logRun()` | Writes entry to RUN_LOG |

### party-ledger.js

Individual party ledger generation:

| Function | Purpose |
|----------|---------|
| `createPartyLedger(ss, party, transactions)` | Creates a party ledger sheet |
| `writeHeader(sheet, party, startRow)` | Writes ledger header (type + code) |
| `writePartyDetails(sheet, party, startRow)` | Writes party info section |
| `writeTransactionTable(sheet, transactions, startRow)` | Writes transaction table |
| `writeTotals(sheet, transactions, startRow)` | Writes totals and closing balance |
| `updateLedgerMasterIndex(ss, parties)` | Updates Ledger Master with all parties |

### utils.js

Helper functions:

| Function | Purpose |
|----------|---------|
| `getColumnIndex(headers, columnName)` | Finds column by name (case-insensitive) |
| `getValue(row, headers, columnName, defaultValue)` | Gets value from row by column name |
| `parseDate(value)` | Safely parses date values |
| `parseNumber(value)` | Parses numbers, removes currency symbols |
| `formatDate(date, format)` | Formats date for display |
| `formatCurrency(value, currency)` | Formats number as currency |
| `chunk(array, size)` | Splits array into chunks |
| `validateRequired(obj, requiredFields)` | Validates required fields |

### logger.js

Logging and notifications:

| Function | Purpose |
|----------|---------|
| `logRun(orgCode, result, startTime)` | Logs completed run to RUN_LOG |
| `logError(orgCode, operation, error)` | Logs error to RUN_LOG |
| `sendErrorEmail(subject, errors)` | Sends email notification |
| `getRecentLogs(limit, orgCode)` | Retrieves recent log entries |
| `cleanupOldLogs(retentionDays)` | Removes old log entries |

---

## Configuration

### CONFIG Tab Structure

The CONFIG tab contains all settings in key-value format:

#### Organization Settings
| Key | Example | Description |
|-----|---------|-------------|
| `ORG_CODE` | CM | Organization code (CM, KM, DEV, CPI, SM) |
| `ORG_NAME` | Congzhou Machinery | Full organization name |
| `FINANCIAL_YEAR` | 2025-26 | Financial year |

#### Source Sheets
| Key | Example | Description |
|-----|---------|-------------|
| `PURCHASE_SHEET_ID` | 1abc...xyz | Google Sheet ID for Purchase Register |
| `PURCHASE_SHEET_NAME` | Purchase Register | Tab name in the sheet |
| `SALES_SHEET_ID` | 1def...uvw | Google Sheet ID for Sales Register |
| `SALES_SHEET_NAME` | Sales Register | Tab name in the sheet |
| `BANK_SHEET_ID` | 1ghi...rst | Google Sheet ID for Bank Statement |
| `BANK_SHEET_NAME` | Bank Statement | Tab name in the sheet |

#### Column Mappings - Purchase
| Key | Default | Description |
|-----|---------|-------------|
| `PUR_DATE_COL` | INW_DATE | Date column |
| `PUR_PARTY_ID_COL` | L/F | Party ID column |
| `PUR_PARTY_NAME_COL` | SUPPLIER_NAME | Party name column |
| `PUR_INVOICE_COL` | INVOICE_NO | Invoice number column |
| `PUR_AMOUNT_COL` | GRAND_TOTAL | Total amount column |

#### Column Mappings - Sales
| Key | Default | Description |
|-----|---------|-------------|
| `SAL_DATE_COL` | INVOICE_DATE | Date column |
| `SAL_PARTY_ID_COL` | L/F | Party ID column |
| `SAL_PARTY_NAME_COL` | PARTY_NAME | Party name column |
| `SAL_INVOICE_COL` | INVOICE_NO | Invoice number column |
| `SAL_AMOUNT_COL` | GRAND_TOTAL | Total amount column |

#### Column Mappings - Bank
| Key | Default | Description |
|-----|---------|-------------|
| `BANK_DATE_COL` | DATE | Date column |
| `BANK_PARTY_ID_COL` | L/F | Party ID column |
| `BANK_DEBIT_COL` | DEBIT | Debit column |
| `BANK_CREDIT_COL` | CREDIT | Credit column |
| `BANK_PARTICULARS_COL` | PARTICULARS | Particulars column |

#### Notifications
| Key | Default | Description |
|-----|---------|-------------|
| `ERROR_EMAIL` | | Email for error notifications |
| `LOG_RETENTION_DAYS` | 30 | Days to keep run logs |

---

## Party Ledger Format

Individual party ledger sheets follow this structure:

```
┌─────────────────────────────────────────────────────────────────────┐
│ SUPPLIER LEDGER                              CG-SUP-0001            │
├─────────────────────────────────────────────────────────────────────┤
│ PARENT COMPANY NAME     [Parent Company]                            │
│ ADDRESS LINE 1          [Address]                                   │
│ ADDRESS LINE 2          [City, State]                               │
│ CONTACT                 [Phone/Email]                               │
├─────────────────────────────────────────────────────────────────────┤
│ PARTY NAME              [Party Name]                                │
│ ADDRESS LINE 1          [Address]                                   │
│ ADDRESS LINE 2          [City, State]                               │
│ GST NUMBER              [GSTIN]                                     │
├─────────────────────────────────────────────────────────────────────┤
│ DATE     │ PARTICULARS          │ VOUCHER │ REF │ DEBIT  │ CREDIT  │
│ 15-01-26 │ Purchase Invoice 123 │ PURCHASE│ 123 │        │ 54,000  │
│ 20-01-26 │ Payment via NEFT     │ BANK    │ 456 │ 50,000 │         │
│ ...      │ ...                  │ ...     │ ... │ ...    │ ...     │
├─────────────────────────────────────────────────────────────────────┤
│                                TOTAL:          │ 50,000 │ 54,000   │
│ CLOSING BALANCE                                │        │ 4,000 CR │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Data Flow

```
┌──────────────────┐     ┌──────────────────┐     ┌──────────────────┐
│ Purchase Register│     │ Sales Register   │     │ Bank Statement   │
│ (Source Sheet)   │     │ (Source Sheet)   │     │ (Source Sheet)   │
└────────┬─────────┘     └────────┬─────────┘     └────────┬─────────┘
         │                        │                        │
         │    fetchSourceData()   │                        │
         ▼                        ▼                        ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      transformRow()                                  │
│    Standardizes to: {date, partyId, partyName, debit, credit, ...}  │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│                   generateAllLedgers()                               │
│                                                                      │
│   1. Group transactions by partyId                                   │
│   2. Calculate totals per party                                      │
│   3. Create individual party ledger sheets                           │
│   4. Update Ledger Master index                                      │
└─────────────────────────────────────────────────────────────────────┘
                              │
         ┌────────────────────┼────────────────────┐
         ▼                    ▼                    ▼
┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
│ Ledger Master   │  │ CG-SUP-0001     │  │ CG-CUST-0001    │
│ (Index)         │  │ (Party Ledger)  │  │ (Party Ledger)  │
└─────────────────┘  └─────────────────┘  └─────────────────┘
```

---

## Deployment

### Development Workflow

```bash
# Navigate to project
cd ~/Desktop/development/accounts-script

# Make changes to src/ files

# Push to Google
npx @google/clasp push

# Open in browser
npx @google/clasp open

# View logs
npx @google/clasp logs
```

### Test as Add-on

1. Open Apps Script editor: `npx @google/clasp open`
2. Click **Deploy → Test deployments**
3. Click **Install** (installs for your account only)
4. Open any Google Sheet
5. Go to **Extensions → CG Accounts Add-on**

### Publish to Workspace

1. Open Apps Script editor
2. Click **Deploy → New deployment**
3. Select **Add-on** as deployment type
4. Fill in:
   - Description: "CG Accounts - Financial data aggregation"
   - Access: "Workspace internal only"
5. Click **Deploy**
6. Copy the deployment ID

### Install Domain-Wide (Admin)

1. Go to [Google Workspace Admin Console](https://admin.google.com)
2. Navigate to **Apps → Google Workspace Marketplace apps**
3. Click **Add internal app**
4. Paste the deployment ID
5. Configure installation: All users or specific groups
6. Click **Install**

---

## Usage

### First Time Setup (Per Sheet)

1. Open any Google Sheet
2. Go to **Extensions → CG Accounts Add-on → Open Dashboard**
3. Click **Initialize Sheet**
4. Go to **CONFIG** tab and fill in:
   - `ORG_CODE`: Your org code (CM, KM, DEV, CPI, SM)
   - `PURCHASE_SHEET_ID`: ID of your Purchase Register sheet
   - `SALES_SHEET_ID`: ID of your Sales Register sheet
   - `BANK_SHEET_ID`: ID of your Bank Statement sheet
5. Click **Test Connection** to verify
6. Click **Refresh Data** to fetch and generate ledgers

### Regular Usage

1. Open your Ledger sheet
2. Go to **CG Accounts → Refresh Data**
3. View generated ledgers in individual tabs
4. Use **Ledger Master** tab for quick navigation

### Hourly Auto-Refresh

To enable automatic hourly refresh:

1. Open Apps Script editor
2. Go to **Triggers** (clock icon)
3. Click **Add Trigger**
4. Configure:
   - Function: `hourlyRefresh`
   - Event source: Time-driven
   - Type: Hour timer
   - Interval: Every hour
5. Click **Save**

---

## Troubleshooting

### "CONFIG tab not found"

Run **Initialize Sheet** first.

### "Sheet not found: Purchase Register"

Check that:
1. The `PURCHASE_SHEET_ID` is correct
2. The `PURCHASE_SHEET_NAME` matches the tab name exactly

### "Error: You do not have permission"

The source sheet hasn't been shared with you. Ask the owner to share it with your Google account.

### Ledgers not updating

1. Check **RUN_LOG** tab for errors
2. Verify source sheet IDs are correct
3. Run **Test Connection** to verify access

---

## Organizations

| Code | Name | Status |
|------|------|--------|
| CM | Congzhou Machinery (Unipack) | Active |
| KM | Kabira Mobility | Active |
| DEV | DeltaEV | Active |
| CPI | Classic Packaging Industry | Active |
| SM | Satlok Machinery | Future |

---

## Changelog

### v0.2.0 (2026-01-16)
- Changed to standalone add-on architecture
- Added per-sheet CONFIG tab (no central control panel)
- Added party ledger template with header, details, totals
- Added Ledger Master index with hyperlinks
- Added add-on card UI for sidebar
- Added onHomepage trigger for add-on

### v0.1.0 (2026-01-16)
- Initial project setup
- Core modules: fetchers, ledgers, config, logger
- Material Design sidebar UI
- Jest unit tests

---

## Support

For issues or questions:
- Check RUN_LOG tab for error details
- Contact: Sagar Siwach

---

## License

Private - Classic Group internal use only.
