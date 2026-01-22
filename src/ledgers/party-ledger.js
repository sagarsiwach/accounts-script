/**
 * Party Ledger Module
 * Creates individual party ledger sheets with proper formatting
 *
 * @fileoverview Party ledger generation matching the exact template format
 *
 * LEDGER LAYOUT:
 * Row 1: Ledger type (A1:E1 merged) + Contact ID (F1) + Company ID (G1)
 * Row 2: Company Name (A2:F2 merged, 16px, center)
 * Row 3: Company Address Line 1 (A3:F3 merged, 10px, center)
 * Row 4: Company Address Line 2 (A4:F4 merged, 10px, center)
 * Row 5: GST | Phone | Email (A5:F5 merged, 10px, center)
 * Row 6: Empty (20px)
 * Row 7: Party Name (A7:F7 merged, 16px, center)
 * Row 8: Party Address Line 1 (A8:F8 merged, 10px, center)
 * Row 9: Party Address Line 2 (A9:F9 merged, 10px, center)
 * Row 10: GST | Phone | Email (A10:F10 merged, 10px, center)
 * Row 11: Empty
 * Row 12: Section header (merged, 12px, 30px, top/bottom border)
 * Row 13: Legend - DATE, PARTICULARS, VOUCHER TYPE, REF, DEBIT, CREDIT, FILE
 * Row 14+: Transactions (min 10 rows)
 * After transactions: 2 row gap, TOTAL, CLOSING BALANCE, GRAND TOTAL
 */

const PartyLedger = (function() {

  // Currency format for Indian Rupees
  const CURRENCY_FORMAT = '_("₹"* #,##0.00_);_("₹"* \\(#,##0.00\\);_("₹"* "-"??_);_(@_)';

  // Column widths as specified
  const COL_WIDTHS = {
    A: 120,
    B: 350,
    C: 120,
    D: 140,  // Increased by 60
    E: 140,
    F: 140,
    G: 230   // Hidden column for IDs
  };

  /**
   * Creates or updates a party ledger sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   * @param {Object} party - Party data object
   * @param {Object} company - Company data object (the org whose ledger this is)
   * @param {Array} transactions - Array of transaction objects
   * @param {string} ledgerCategory - Ledger category: 'SU' for supplier, 'CU' for customer
   * @returns {Object} Result with sheet name and row count
   */
  function createPartyLedger(ss, party, company, transactions, ledgerCategory) {
    const sheetName = buildSheetName(party.id, ledgerCategory);
    let sheet = ss.getSheetByName(sheetName);

    // Create or clear sheet
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
      sheet.clearFormats();
    }

    // Apply global font
    sheet.getDataRange().setFontFamily('Roboto Condensed');

    // Hide gridlines
    sheet.setHiddenGridlines(true);

    // Set column widths
    sheet.setColumnWidth(1, COL_WIDTHS.A);  // A - Date
    sheet.setColumnWidth(2, COL_WIDTHS.B);  // B - Particulars
    sheet.setColumnWidth(3, COL_WIDTHS.C);  // C - Voucher Type
    sheet.setColumnWidth(4, COL_WIDTHS.D);  // D - Ref
    sheet.setColumnWidth(5, COL_WIDTHS.E);  // E - Debit
    sheet.setColumnWidth(6, COL_WIDTHS.F);  // F - Credit
    sheet.setColumnWidth(7, COL_WIDTHS.G);  // G - Hidden ID column

    // Hide column G and delete columns after G
    sheet.hideColumns(7);
    const maxCols = sheet.getMaxColumns();
    if (maxCols > 7) {
      sheet.deleteColumns(8, maxCols - 7);
    }

    // === ROW 1: Header with ledger type and IDs ===
    writeRow1Header(sheet, party, company, ledgerCategory);

    // === ROWS 2-5: Company Details ===
    writeCompanyDetails(sheet, company);

    // === ROW 6: Empty spacer ===
    sheet.setRowHeight(6, 20);

    // === ROWS 7-10: Party Details ===
    writePartyDetails(sheet, party);

    // === ROW 11: Empty ===
    sheet.setRowHeight(11, 20);

    // === ROW 12: Section header ===
    writeRow12SectionHeader(sheet, party);

    // === ROW 13: Legend ===
    writeRow13Legend(sheet);

    // === ROW 14+: Transactions ===
    const transactionEndRow = writeTransactions(sheet, transactions);

    // === TOTALS SECTION ===
    const grandTotalRow = writeTotalsSection(sheet, transactions, transactionEndRow);

    // === CLEANUP: Delete unused rows after GRAND TOTAL ===
    const maxRows = sheet.getMaxRows();
    if (maxRows > grandTotalRow) {
      sheet.deleteRows(grandTotalRow + 1, maxRows - grandTotalRow);
    }

    // === CLEANUP: Ensure column G hidden and columns after G deleted ===
    sheet.hideColumns(7);
    const finalMaxCols = sheet.getMaxColumns();
    if (finalMaxCols > 7) {
      sheet.deleteColumns(8, finalMaxCols - 7);
    }

    return {
      sheetName: sheetName,
      rowCount: transactions.length,
      ledgerCategory: ledgerCategory
    };
  }

  /**
   * Row 1: Ledger type label + Contact ID + Company ID
   * A1:E1 merged, light gray text, 10px
   * F1: Party contact ID (light gray)
   * G1: Company contact ID (light gray)
   */
  function writeRow1Header(sheet, party, company, ledgerCategory) {
    sheet.setRowHeight(1, 30);

    // Determine ledger type based on category
    let ledgerType = ledgerCategory === 'SU' ? 'SUPPLIER LEDGER' : 'CUSTOMER LEDGER';

    // A1:E1 merged - ledger type
    sheet.getRange('A1:E1').merge()
      .setValue(ledgerType)
      .setFontSize(10)
      .setFontColor('#999999')
      .setVerticalAlignment('middle');

    // F1 - Party ID (light gray)
    sheet.getRange('F1')
      .setValue(party.id || '')
      .setFontSize(10)
      .setFontColor('#999999')
      .setHorizontalAlignment('right')
      .setVerticalAlignment('middle');

    // G1 - Company ID (light gray)
    sheet.getRange('G1')
      .setValue(company.id || '')
      .setFontSize(10)
      .setFontColor('#999999')
      .setHorizontalAlignment('right')
      .setVerticalAlignment('middle');
  }

  /**
   * Rows 2-5: Company details (the org running the ledger)
   */
  function writeCompanyDetails(sheet, company) {
    // Row 2: Company Name - 16px, center, 30px height, UPPERCASE, BOLD
    sheet.setRowHeight(2, 30);
    sheet.getRange('A2:F2').merge()
      .setValue((company.name || '').toUpperCase())
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Row 3: Address Line 1 - 10px, center, 22px height
    sheet.setRowHeight(3, 22);
    sheet.getRange('A3:F3').merge()
      .setValue(company.address1 || '')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Row 4: Address Line 2 - 10px, center, 22px height
    sheet.setRowHeight(4, 22);
    sheet.getRange('A4:F4').merge()
      .setValue(company.address2 || '')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Row 5: GST | Phone | Email - 10px, center, 22px height
    sheet.setRowHeight(5, 22);
    const contactInfo = buildContactInfo(company.gst, company.phone, company.email);
    sheet.getRange('A5:F5').merge()
      .setValue(contactInfo)
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  /**
   * Rows 7-10: Party details (supplier/customer whose ledger this is)
   */
  function writePartyDetails(sheet, party) {
    // Row 7: Party Name - 16px, center, 30px height, UPPERCASE, BOLD
    sheet.setRowHeight(7, 30);
    sheet.getRange('A7:F7').merge()
      .setValue((party.name || '').toUpperCase())
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Row 8: Address Line 1 - 10px, center, 22px height
    sheet.setRowHeight(8, 22);
    sheet.getRange('A8:F8').merge()
      .setValue(party.address1 || '')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Row 9: Address Line 2 - 10px, center, 22px height
    sheet.setRowHeight(9, 22);
    sheet.getRange('A9:F9').merge()
      .setValue(party.address2 || '')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Row 10: GST | Phone | Email - 10px, center, 22px height
    sheet.setRowHeight(10, 22);
    const contactInfo = buildContactInfo(party.gst, party.phone, party.email);
    sheet.getRange('A10:F10').merge()
      .setValue(contactInfo)
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  /**
   * Build contact info string with pipe separators
   * Only includes non-empty values
   */
  function buildContactInfo(gst, phone, email) {
    const parts = [];
    if (gst) parts.push(gst);
    if (phone) parts.push(phone);
    if (email) parts.push(email);
    return parts.join(' | ');
  }

  /**
   * Row 12: Section header (empty merged row with borders)
   */
  function writeRow12SectionHeader(sheet, party) {
    sheet.setRowHeight(12, 30);
    sheet.getRange('A12:G12').merge()
      .setFontSize(12)
      .setBorder(true, null, true, null, null, null); // top and bottom borders
  }

  /**
   * Row 13: Legend row with column headers
   */
  function writeRow13Legend(sheet) {
    sheet.setRowHeight(13, 25);

    const headers = ['DATE', 'PARTICULARS', 'VOUCHER TYPE', 'REF', 'DEBIT', 'CREDIT', ''];
    sheet.getRange('A13:G13').setValues([headers])
      .setFontSize(10)
      .setFontWeight('bold')
      .setBorder(true, null, true, null, null, null) // top and bottom borders
      .setVerticalAlignment('middle');

    // Alignment
    sheet.getRange('A13:B13').setHorizontalAlignment('left');
    sheet.getRange('C13:D13').setHorizontalAlignment('center');
    sheet.getRange('E13:F13').setHorizontalAlignment('right');
  }

  /**
   * Write transaction rows starting from row 14
   * Minimum 10 rows even if fewer transactions
   * Returns the last row number of the transaction section
   */
  function writeTransactions(sheet, transactions) {
    const startRow = 14;
    const minRows = 10;
    const actualRows = Math.max(transactions.length, minRows);

    // Set row heights for all transaction rows
    for (let i = 0; i < actualRows; i++) {
      sheet.setRowHeight(startRow + i, 22);
    }

    // Write transaction data
    if (transactions.length > 0) {
      const dataRows = transactions.map(txn => [
        txn.date || '',
        txn.particulars || txn.narration || '',
        txn.voucherType || txn.docType || '',
        txn.reference || txn.docNo || '',
        txn.debit || '',
        txn.credit || '',
        '' // File column - leave empty for now
      ]);

      sheet.getRange(startRow, 1, dataRows.length, 7).setValues(dataRows);
    }

    // Format all transaction rows
    const range = sheet.getRange(startRow, 1, actualRows, 7);
    range.setFontSize(9);

    // Alignment: A, B left; C, D center; E, F currency
    sheet.getRange(startRow, 1, actualRows, 2).setHorizontalAlignment('left');
    sheet.getRange(startRow, 3, actualRows, 2).setHorizontalAlignment('center');
    sheet.getRange(startRow, 5, actualRows, 2)
      .setHorizontalAlignment('right')
      .setNumberFormat(CURRENCY_FORMAT);

    // Format date column
    if (transactions.length > 0) {
      sheet.getRange(startRow, 1, transactions.length, 1).setNumberFormat('dd-mm-yyyy');
    }

    return startRow + actualRows - 1;
  }

  /**
   * Write totals section after transactions
   * 2 row gap, then TOTAL, CLOSING BALANCE, GRAND TOTAL
   */
  function writeTotalsSection(sheet, transactions, lastTransactionRow) {
    // Calculate totals
    let totalDebit = 0;
    let totalCredit = 0;

    for (const txn of transactions) {
      totalDebit += parseFloat(txn.debit) || 0;
      totalCredit += parseFloat(txn.credit) || 0;
    }

    const closingBalance = totalDebit - totalCredit;

    // 2 row gap
    const gapRow1 = lastTransactionRow + 1;
    const gapRow2 = lastTransactionRow + 2;
    sheet.setRowHeight(gapRow1, 20);
    sheet.setRowHeight(gapRow2, 20);

    // TOTAL row - only top border, bold
    const totalRow = lastTransactionRow + 3;
    sheet.setRowHeight(totalRow, 25);
    sheet.getRange(totalRow, 2).setValue('TOTAL')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    sheet.getRange(totalRow, 5).setValue(totalDebit)
      .setNumberFormat(CURRENCY_FORMAT)
      .setFontWeight('bold');
    sheet.getRange(totalRow, 6).setValue(totalCredit)
      .setNumberFormat(CURRENCY_FORMAT)
      .setFontWeight('bold');
    sheet.getRange(totalRow, 1, 1, 7).setBorder(true, null, null, null, null, null); // top border only

    // CLOSING BALANCE row - no border, not bold
    const closingRow = totalRow + 1;
    sheet.setRowHeight(closingRow, 25);
    sheet.getRange(closingRow, 2).setValue('CLOSING BALANCE')
      .setHorizontalAlignment('right');

    // Put closing balance in debit or credit column based on sign
    if (closingBalance >= 0) {
      sheet.getRange(closingRow, 5).setValue(Math.abs(closingBalance))
        .setNumberFormat(CURRENCY_FORMAT);
      sheet.getRange(closingRow, 7).setValue('DR');
    } else {
      sheet.getRange(closingRow, 6).setValue(Math.abs(closingBalance))
        .setNumberFormat(CURRENCY_FORMAT);
      sheet.getRange(closingRow, 7).setValue('CR');
    }

    // GRAND TOTAL row - bold, top and bottom border
    // Grand total = Total + Closing Balance adjustment (so Debit = Credit after adjustment)
    const grandTotalRow = closingRow + 1;
    sheet.setRowHeight(grandTotalRow, 25);
    sheet.getRange(grandTotalRow, 2).setValue('GRAND TOTAL')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');

    // Grand totals balance out (Dr = Cr after closing balance)
    const maxTotal = Math.max(totalDebit, totalCredit);
    sheet.getRange(grandTotalRow, 5).setValue(maxTotal)
      .setNumberFormat(CURRENCY_FORMAT)
      .setFontWeight('bold');
    sheet.getRange(grandTotalRow, 6).setValue(maxTotal)
      .setNumberFormat(CURRENCY_FORMAT)
      .setFontWeight('bold');
    sheet.getRange(grandTotalRow, 1, 1, 7).setBorder(true, null, true, null, null, null); // top and bottom

    // Apply font to entire sheet
    sheet.getDataRange().setFontFamily('Roboto Condensed');

    return grandTotalRow;
  }

  /**
   * Builds sheet name in format: [SU] CG-SUP-0001 or [CU] CG-CUS-0001
   * @param {string} partyId - Party ID (e.g., CG-SUP-0001)
   * @param {string} ledgerCategory - 'SU' or 'CU'
   * @returns {string} Formatted sheet name
   */
  function buildSheetName(partyId, ledgerCategory) {
    const prefix = '[' + (ledgerCategory || 'CU') + '] ';
    let id = String(partyId || 'Unknown')
      .replace(/[\/\\?*\[\]:]/g, '-')
      .trim();

    return prefix + id;
  }

  /**
   * Updates the Ledger Master index with all parties
   * @param {Spreadsheet} ss - Active spreadsheet
   * @param {Array} ledgers - Array of ledger summary objects (with ledgerCategory)
   */
  function updateLedgerMasterIndex(ss, ledgers) {
    const sheet = ss.getSheetByName('Ledger Master');
    if (!sheet) return;

    // Clear existing data (keep header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 8).clear();
    }

    if (ledgers.length === 0) {
      sheet.getRange(2, 1, 1, 8).merge()
        .setValue('No ledgers found. Run Refresh to generate.')
        .setFontStyle('italic')
        .setFontColor('#666666')
        .setHorizontalAlignment('center');
      return;
    }

    // Sort ledgers by category (SU first, then CU) then by name
    ledgers.sort((a, b) => {
      // SU before CU
      if (a.ledgerCategory !== b.ledgerCategory) {
        return a.ledgerCategory === 'SU' ? -1 : 1;
      }
      return (a.name || '').localeCompare(b.name || '');
    });

    // Build rows with hyperlinks
    const rows = ledgers.map(ledger => {
      const sheetName = buildSheetName(ledger.id, ledger.ledgerCategory);
      const link = '=HYPERLINK("#gid=' + getSheetGid(ss, sheetName) + '", "Open")';

      return [
        ledger.id,
        ledger.name,
        '[' + ledger.ledgerCategory + ']',  // Show category as [SU] or [CU]
        ledger.totalDebit || 0,
        ledger.totalCredit || 0,
        (ledger.totalDebit || 0) - (ledger.totalCredit || 0),
        ledger.lastTransaction || '',
        link
      ];
    });

    // Write data
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);

    // Format currency columns
    sheet.getRange(2, 4, rows.length, 3).setNumberFormat(CURRENCY_FORMAT);

    // Format date column
    sheet.getRange(2, 7, rows.length, 1).setNumberFormat('dd-mm-yyyy');

    // Alternate row colors
    for (let i = 0; i < rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(2 + i, 1, 1, 8).setBackground('#f8f9fa');
      }
    }

    // Apply Roboto Condensed font
    sheet.getDataRange().setFontFamily('Roboto Condensed');
  }

  /**
   * Gets the GID of a sheet for hyperlink
   */
  function getSheetGid(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return '0';
    return sheet.getSheetId();
  }

  /**
   * Sanitizes a string for use as a sheet name
   */
  function sanitizeSheetName(name) {
    let sanitized = String(name)
      .replace(/[\/\\?*\[\]:]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();

    if (sanitized.length > 100) {
      sanitized = sanitized.substring(0, 97) + '...';
    }

    return sanitized || 'Unnamed';
  }

  // Public API
  return {
    createPartyLedger,
    updateLedgerMasterIndex,
    sanitizeSheetName,
    buildSheetName
  };

})();
