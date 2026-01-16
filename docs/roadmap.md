# Classic Group Accounts Script - Roadmap

**Project Start:** 2026-01-16
**Target Completion:** 2026-02-15 (4 weeks)

---

## Project Overview

### Purpose
Automated financial data aggregation system for Classic Group entities. Pulls from Purchase, Sales, and Bank source sheets, consolidates into Ledger Masters, and generates individual party ledgers.

### Organizations Covered

| Code | Name | Priority |
|------|------|----------|
| CM | Congzhou Machinery (Unipack) | High |
| KM | Kabira Mobility | High |
| DEV | DeltaEV | High |
| CPI | Classic Packaging Industry | Medium |
| SM | Satlok Machinery | Low |

### Users
- Accounts team at classicgroup.asia (Google Workspace)
- Primary: Achal, Sagar

### Key Features
1. **Central Control Panel** - One config sheet controls all orgs
2. **Automated Data Fetch** - Pulls from Purchase, Sales, Bank registers
3. **Ledger Generation** - Creates Supplier, Customer, Contractor ledgers
4. **Professional UI** - Material Design sidebar with clean, minimal interface
5. **Scheduled Execution** - Hourly trigger with manual override
6. **Error Handling** - Logging + email notifications
7. **Multi-org Support** - Add new org without code changes

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    CG ACCOUNTS CONTROL                          │
│                  (Central Config Sheet)                         │
├─────────────────────────────────────────────────────────────────┤
│ ORG_CONFIG | SOURCE_CONFIG | COLUMN_MAPPING | RUN_LOG          │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐     ┌───────────────┐     ┌───────────────┐
│ CM Sources    │     │ KM Sources    │     │ DEV Sources   │
│ - Purchase    │     │ - Purchase    │     │ - Purchase    │
│ - Sales       │     │ - Sales       │     │ - Sales       │
│ - Bank        │     │ - Bank        │     │ - Bank        │
└───────┬───────┘     └───────┬───────┘     └───────┬───────┘
        │                     │                     │
        ▼                     ▼                     ▼
┌───────────────┐     ┌───────────────┐     ┌───────────────┐
│ CM Ledger     │     │ KM Ledger     │     │ DEV Ledger    │
│ Master        │     │ Master        │     │ Master        │
├───────────────┤     ├───────────────┤     ├───────────────┤
│ - Supplier    │     │ - Supplier    │     │ - Supplier    │
│ - Customer    │     │ - Customer    │     │ - Customer    │
│ - Contractor  │     │ - Contractor  │     │ - Contractor  │
└───────────────┘     └───────────────┘     └───────────────┘
```

---

## Phase Breakdown

### Phase 1: Foundation (Week 1)
**Target: 2026-01-23**

| Task | Status | Notes |
|------|--------|-------|
| Project setup (Git, clasp, folder structure) | Done | |
| Create roadmap and session logging | In Progress | |
| package.json with Jest config | Pending | |
| appsscript.json manifest | Pending | |
| Config module (read ORG_CONFIG, SOURCE_CONFIG) | Pending | |
| Logger module (sheets + email) | Pending | |
| Unit tests setup | Pending | |

### Phase 2: Data Fetchers (Week 2)
**Target: 2026-01-30**

| Task | Status | Notes |
|------|--------|-------|
| Purchase fetcher | Pending | |
| Sales fetcher | Pending | |
| Bank fetcher | Pending | |
| Data transformation layer | Pending | |
| Unit tests for fetchers | Pending | |

### Phase 3: Ledger Generation (Week 3)
**Target: 2026-02-06**

| Task | Status | Notes |
|------|--------|-------|
| Ledger Master generation | Pending | |
| Supplier Ledger generation | Pending | |
| Customer Ledger generation | Pending | |
| Contractor Ledger generation | Pending | |
| Unit tests for ledgers | Pending | |

### Phase 4: UI & Polish (Week 4)
**Target: 2026-02-13**

| Task | Status | Notes |
|------|--------|-------|
| Sidebar HTML (Material Design) | Pending | |
| Menu integration | Pending | |
| Refresh All functionality | Pending | |
| Refresh Single Org functionality | Pending | |
| View Logs panel | Pending | |
| Settings panel | Pending | |
| Error email notifications | Pending | |
| Hourly trigger setup | Pending | |
| Final testing | Pending | |
| Deployment to classicgroup.asia | Pending | |

---

## Technical Stack

| Component | Technology |
|-----------|------------|
| Runtime | Google Apps Script (V8) |
| Local Dev | clasp + Node.js |
| Testing | Jest |
| UI | HTML Service + Material Design Lite |
| Version Control | Git + GitHub |
| Deployment | clasp push |

---

## Config Sheet Structure

### ORG_CONFIG
| Column | Type | Description |
|--------|------|-------------|
| ORG_CODE | Text | CM, KM, DEV, CPI, SM |
| ORG_NAME | Text | Full organization name |
| LEDGER_SHEET_ID | Text | Google Sheet ID for ledger output |
| ACTIVE | Boolean | Whether to process this org |

### SOURCE_CONFIG
| Column | Type | Description |
|--------|------|-------------|
| ORG_CODE | Text | Which org this belongs to |
| SOURCE_TYPE | Text | PURCHASE, SALES, BANK |
| SHEET_ID | Text | Source Google Sheet ID |
| SHEET_NAME | Text | Tab name within the sheet |
| ACTIVE | Boolean | Whether to fetch from this source |

### COLUMN_MAPPING
| Column | Type | Description |
|--------|------|-------------|
| SOURCE_TYPE | Text | PURCHASE, SALES, BANK |
| FIELD | Text | Standardized field name (DATE, PARTY_ID, etc.) |
| COLUMN_NAME | Text | Actual column name in source |
| REQUIRED | Boolean | Is this field required? |

### RUN_LOG
| Column | Type | Description |
|--------|------|-------------|
| TIMESTAMP | DateTime | When the run happened |
| ORG_CODE | Text | Which org was processed |
| SOURCE_TYPE | Text | Which source was fetched |
| ROWS_FETCHED | Number | How many rows retrieved |
| ROWS_WRITTEN | Number | How many rows written to ledger |
| STATUS | Text | SUCCESS, ERROR, SKIPPED |
| ERROR_MESSAGE | Text | Error details if failed |
| DURATION_MS | Number | How long it took |

---

## Ledger Formats

### Ledger Master (One per org)
Standard double-entry format:
```
DATE | DOC_TYPE | DOC_NO | ACCOUNT_CODE | PARTY_ID | PARTY_NAME | DEBIT | CREDIT | BALANCE | NARRATION | SOURCE_REF
```

### Party Ledgers (Supplier/Customer/Contractor)
Each party gets a separate sheet within the Ledger workbook:
```
DATE | DOC_TYPE | DOC_NO | DEBIT | CREDIT | BALANCE | NARRATION
```

Opening balance row at top, running balance calculated.

---

## Error Handling

1. **Logging**: All runs logged to RUN_LOG sheet
2. **Email Alerts**: Sent on ERROR status to configured recipients
3. **Graceful Degradation**: One org failing doesn't stop others

---

## Success Metrics

| Metric | Target |
|--------|--------|
| Ledger refresh time (per org) | < 30 seconds |
| Full refresh (all orgs) | < 3 minutes |
| Uptime | 99% (hourly runs) |
| Error rate | < 5% of runs |

---

## Risks & Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Source sheet structure changes | Ledger generation fails | Column mapping config allows easy updates |
| Google quota limits | Script timeout | Batch processing, efficient API calls |
| Permission issues | Can't read source sheets | Clear setup docs, service account option |

---

## Session Log

All development sessions logged in `/docs/sessions/`:
- `session-001.md` - 2026-01-16 - Project setup

---

*Last Updated: 2026-01-16*
