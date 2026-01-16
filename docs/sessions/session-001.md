# Session 001 - Project Setup

**Date:** 2026-01-16
**Duration:** ~1 hour
**Focus:** Initial project setup and architecture decisions

---

## Context

Starting the Classic Group Accounts Script project. This will aggregate financial data from Purchase, Sales, and Bank registers across 5 organizations into consolidated Ledger Masters, then generate individual party ledgers (Supplier, Customer, Contractor).

## Decisions Made

### Architecture

1. **Central Config Sheet (Option B)** - One master config sheet controls all orgs, rather than config embedded in each org's ledger
2. **Script-based aggregation** - Using Apps Script instead of IMPORTRANGE formulas for:
   - Better performance with large datasets
   - Proper error handling
   - Easier multi-org replication
   - Logging and monitoring

### Technical

1. **Trigger**: Hourly automated refresh
2. **Error handling**: Log to sheet + email notifications
3. **UI**: Material Design sidebar (clean, minimal, professional)
4. **Testing**: Jest for unit tests
5. **Source structure**: All orgs use identical Purchase/Sales/Bank formats

### Scope

1. **FY 2025-26 only** - No historical data migration
2. **5 Organizations**: CM, KM, DEV, CPI, SM
3. **Users**: Accounts team at classicgroup.asia

## Work Completed

- [x] Created GitHub repo: https://github.com/sagarsiwach/accounts-script
- [x] Initialized clasp for Apps Script development
- [x] Created folder structure:
  ```
  accounts-script/
  ├── src/
  │   ├── fetchers/
  │   ├── ledgers/
  │   └── ui/
  ├── tests/
  │   ├── unit/
  │   └── mocks/
  └── docs/
      └── sessions/
  ```
- [x] Created roadmap.md with full project scope
- [x] Created this session log

## Questions Resolved

| Question | Answer |
|----------|--------|
| Where does config live? | Central sheet (Option B) |
| Trigger frequency? | Hourly |
| Individual ledgers by script? | Yes |
| Who uses this? | Accounts team |
| Google account? | classicgroup.asia (Workspace) |
| Source sheet consistency? | Identical across all orgs |
| Error handling? | Log + Email |
| Historical data? | No, FY 25-26 only |
| Testing? | Jest unit tests |

## Next Steps

1. Set up package.json with Jest config
2. Create appsscript.json manifest
3. Build config module to read ORG_CONFIG and SOURCE_CONFIG
4. Build logger module
5. User to provide ledger template sheet for format reference

## Blockers

- Need ledger template sheet from user to finalize ledger output format

## Notes

- Sidebar actions: Refresh All, Refresh Single Org, View Logs, Settings, Manual Entry
- Party ledgers (Supplier/Customer/Contractor) will be separate sheets within the Ledger workbook to avoid 300+ tabs problem

---

*Session logged by Claude Code*
