# Session 003 - Ledger Separation & Formatting Fixes

**Date:** 2025-01-23
**Duration:** ~1.5 hours
**Focus:** [SU]/[CU] ledger separation, formatting fixes, amount parsing

---

## Context

This session addressed major architectural changes to separate supplier and customer ledgers, fix formatting issues, and resolve amount parsing bugs for Indian number format.

## Changes Made

### 1. Ledger Category Separation (`main.js`)

Introduced `[SU]` Supplier and `[CU]` Customer ledger categories:

| Source | Party Types | Ledger Category |
|--------|-------------|-----------------|
| Purchase Sheet | All | `[SU]` Supplier |
| Sales Sheet | All | `[CU]` Customer |
| Bank Sheet | `-SUP-` | `[SU]` Supplier |
| Bank Sheet | `-CUS-`, `-REN-`, `-DEA-` | `[CU]` Customer |
| Bank Sheet | `-CON-` | Skipped |
| Bank Sheet | `-MAS-` | Based on debit/credit |

**New function:** `determineLedgerCategory(partyId, docType, txn)`
- Routes transactions to correct ledger type
- Master parties (MAS) get dual-ledger: Purchase→SU, Sales→CU, Bank→by payment direction

### 2. Sheet Naming Format (`party-ledger.js`)

| Before | After |
|--------|-------|
| `CG-SUP-0001` | `[SU] CG-SUP-0001` |
| `CG-CUS-0001` | `[CU] CG-CUS-0001` |

**New function:** `buildSheetName(partyId, ledgerCategory)`
- Creates prefixed sheet names with category indicator
- Shorter format using party ID instead of name

### 3. Formatting Fixes (`party-ledger.js`)

| Issue | Fix |
|-------|-----|
| Unused rows remaining | Delete rows after GRAND TOTAL |
| Column G not hidden | Hide column G at end of creation |
| Extra columns | Delete columns after G at end |

Moved cleanup operations to END of `createPartyLedger()` for reliability.

### 4. Amount Parsing Fix (`main.js`)

**Problem:** Indian number format `2,00,000.00` was parsed as `2.00` instead of `200000.00`

**Solution:** Added `parseNumericValue()` function:
```javascript
function parseNumericValue(val) {
  if (!val) return 0;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/[₹$,\s]/g, '');
  return parseFloat(cleaned) || 0;
}
```

Applied to `fetchBankDataAllTabs()` for bank debit/credit columns.

### 5. Sidebar Improvements

| Change | Details |
|--------|---------|
| Version display | `v0.6.0 (Build 10)` in footer |
| Loading overlay fix | Removed `ui.showSidebar()` that was blocking UI during refresh |

### 6. Documentation

Updated `CLAUDE.md` with direct Marketplace SDK link:
```
https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk?project=cg-accounts-add-on
```

### 7. Version Bump

- Version: **0.5.0** → **0.6.0**
- Apps Script versions created: **10**, **11**, **12**

## Files Modified

| File | Changes |
|------|---------|
| `src/main.js` | Ledger separation logic, parseNumericValue(), removed sidebar blocking |
| `src/ledgers/party-ledger.js` | buildSheetName(), cleanup at end, delete unused rows |
| `src/ui/Sidebar.html` | Version display v0.6.0 (Build 10) |
| `CLAUDE.md` | Added direct Marketplace SDK link |

## Commits

1. `2ea2495` - feat: Separate [SU] Supplier and [CU] Customer ledgers
2. `31c2274` - fix: Amount parsing and sheet naming improvements
3. `f0663c4` - docs: Add direct Marketplace SDK link for quick version updates
4. `fca105d` - fix: Remove sidebar UI blocking during data refresh

## Features Discussed (Not Yet Implemented)

| Feature | Description | Priority |
|---------|-------------|----------|
| Opening Balance | Auto-calculate from transactions before FROM_DATE | High |
| Date Range Filter | FROM_DATE/TO_DATE in CONFIG | High |
| Unknown Party Flagging | Log transactions with invalid party IDs | High |
| Source Tracking | Store source sheet in column G | Medium |
| Duplicate Detection | Flag same date+amount+party | Medium |
| Period Subtotals | Monthly/quarterly subtotals | Medium |

## Open Questions

1. **Rental parties (REN):** Should they route to Supplier ledger if you PAY them rent? Currently routes to Customer.
2. **Opening Balance approach:** Auto-calculate from history vs manual entry sheet

## Deployment Workflow

```bash
# From src/ directory
clasp push -f && clasp version "Description"

# Then update Marketplace SDK version at:
# https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk?project=cg-accounts-add-on

# Then git commit
git add -A && git commit -m "message" && git push origin master
```

---

*Session logged by Claude Code*
