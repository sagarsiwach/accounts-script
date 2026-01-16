/**
 * Party Ledger Module
 * Creates individual party ledger sheets with proper formatting
 *
 * @fileoverview Party ledger generation matching the template format
 */

const PartyLedger = (function() {

  /**
   * Creates or updates a party ledger sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   * @param {Object} party - Party data object
   * @param {Array} transactions - Array of transaction objects
   * @returns {Object} Result with sheet name and row count
   */
  function createPartyLedger(ss, party, transactions) {
    const sheetName = sanitizeSheetName(party.id);
    let sheet = ss.getSheetByName(sheetName);

    // Create or clear sheet
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }

    // Build the ledger
    let currentRow = 1;

    // === HEADER SECTION ===
    currentRow = writeHeader(sheet, party, currentRow);

    // === PARTY DETAILS SECTION ===
    currentRow = writePartyDetails(sheet, party, currentRow);

    // === TRANSACTION TABLE ===
    currentRow = writeTransactionTable(sheet, transactions, currentRow);

    // === TOTALS AND CLOSING BALANCE ===
    currentRow = writeTotals(sheet, transactions, currentRow);

    // Apply styling
    applyLedgerStyling(sheet);

    return {
      sheetName: sheetName,
      rowCount: transactions.length
    };
  }

  /**
   * Writes the ledger header (type + code)
   */
  function writeHeader(sheet, party, startRow) {
    const ledgerType = party.type === 'SUPPLIER' ? 'SUPPLIER LEDGER' :
                       party.type === 'CUSTOMER' ? 'CUSTOMER LEDGER' :
                       'CONTRACTOR LEDGER';

    // Merge cells for header
    sheet.getRange(startRow, 1, 1, 4).merge();
    sheet.getRange(startRow, 5, 1, 3).merge();

    sheet.getRange(startRow, 1).setValue(ledgerType)
      .setFontWeight('bold')
      .setFontSize(14);

    sheet.getRange(startRow, 5).setValue(party.id)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('right');

    return startRow + 2;
  }

  /**
   * Writes party details section (parent company, address, contact)
   */
  function writePartyDetails(sheet, party, startRow) {
    const details = [
      ['PARENT COMPANY NAME', party.parentCompany || ''],
      ['ADDRESS LINE 1', party.address1 || ''],
      ['ADDRESS LINE 2', party.address2 || ''],
      ['CONTACT NUMBER / E-MAIL ADDRESS', party.contact || ''],
      ['', ''],
      ['PARTY NAME', party.name || ''],
      ['ADDRESS LINE 1', party.partyAddress1 || party.address1 || ''],
      ['ADDRESS LINE 2', party.partyAddress2 || party.address2 || ''],
      ['GST NUMBER / CONTACT NUMBER / E-MAIL ADDRESS', party.gstNumber || ''],
      ['', '']
    ];

    for (let i = 0; i < details.length; i++) {
      sheet.getRange(startRow + i, 1).setValue(details[i][0])
        .setFontWeight('bold')
        .setFontColor('#666666');
      sheet.getRange(startRow + i, 2, 1, 6).merge();
      sheet.getRange(startRow + i, 2).setValue(details[i][1]);
    }

    return startRow + details.length + 1;
  }

  /**
   * Writes the transaction table header and data
   */
  function writeTransactionTable(sheet, transactions, startRow) {
    // Table header
    const headers = ['DATE', 'PARTICULARS', 'VOUCHER TYPE', 'REF', 'DEBIT', 'CREDIT', 'FILE'];

    sheet.getRange(startRow, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setHorizontalAlignment('center');

    // Transaction rows
    if (transactions.length === 0) {
      sheet.getRange(startRow + 1, 1, 1, headers.length).merge()
        .setValue('No transactions found')
        .setFontStyle('italic')
        .setFontColor('#999999')
        .setHorizontalAlignment('center');
      return startRow + 3;
    }

    const dataRows = transactions.map(txn => [
      txn.date,
      txn.particulars || txn.narration || '',
      txn.voucherType || txn.docType || '',
      txn.reference || txn.docNo || '',
      txn.debit || '',
      txn.credit || '',
      txn.fileLink || ''
    ]);

    sheet.getRange(startRow + 1, 1, dataRows.length, headers.length).setValues(dataRows);

    // Format date column
    sheet.getRange(startRow + 1, 1, dataRows.length, 1).setNumberFormat('dd-mm-yyyy');

    // Format currency columns
    sheet.getRange(startRow + 1, 5, dataRows.length, 2).setNumberFormat('₹ #,##0.00');

    return startRow + 1 + dataRows.length + 1;
  }

  /**
   * Writes totals and closing balance
   */
  function writeTotals(sheet, transactions, startRow) {
    // Calculate totals
    let totalDebit = 0;
    let totalCredit = 0;

    for (const txn of transactions) {
      totalDebit += parseFloat(txn.debit) || 0;
      totalCredit += parseFloat(txn.credit) || 0;
    }

    const closingBalance = totalDebit - totalCredit;

    // Total row
    sheet.getRange(startRow, 1, 1, 4).merge();
    sheet.getRange(startRow, 1).setValue('')
      .setBackground('#f3f3f3');

    sheet.getRange(startRow, 5).setValue(totalDebit)
      .setNumberFormat('₹ #,##0.00')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    sheet.getRange(startRow, 6).setValue(totalCredit)
      .setNumberFormat('₹ #,##0.00')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    sheet.getRange(startRow, 7).setValue('')
      .setBackground('#f3f3f3');

    // Closing balance row
    startRow++;
    sheet.getRange(startRow, 1, 1, 4).merge();
    sheet.getRange(startRow, 1).setValue('CLOSING BALANCE')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#d9ead3');

    sheet.getRange(startRow, 5).setValue('')
      .setBackground('#d9ead3');

    sheet.getRange(startRow, 6).setValue(Math.abs(closingBalance))
      .setNumberFormat('₹ #,##0.00')
      .setFontWeight('bold')
      .setBackground('#d9ead3');

    sheet.getRange(startRow, 7).setValue(closingBalance >= 0 ? 'DR' : 'CR')
      .setFontWeight('bold')
      .setBackground('#d9ead3');

    return startRow + 1;
  }

  /**
   * Applies consistent styling to the ledger sheet
   */
  function applyLedgerStyling(sheet) {
    // Set column widths
    sheet.setColumnWidth(1, 100);  // Date
    sheet.setColumnWidth(2, 350);  // Particulars
    sheet.setColumnWidth(3, 120);  // Voucher Type
    sheet.setColumnWidth(4, 80);   // Ref
    sheet.setColumnWidth(5, 120);  // Debit
    sheet.setColumnWidth(6, 120);  // Credit
    sheet.setColumnWidth(7, 200);  // File

    // Set default font
    sheet.getDataRange().setFontFamily('Arial');
  }

  /**
   * Updates the Ledger Master index with all parties
   * @param {Spreadsheet} ss - Active spreadsheet
   * @param {Array} parties - Array of party summary objects
   */
  function updateLedgerMasterIndex(ss, parties) {
    const sheet = ss.getSheetByName('Ledger Master');
    if (!sheet) return;

    // Clear existing data (keep header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 8).clear();
    }

    if (parties.length === 0) {
      sheet.getRange(2, 1, 1, 8).merge()
        .setValue('No ledgers found. Run Refresh to generate.')
        .setFontStyle('italic')
        .setFontColor('#666666')
        .setHorizontalAlignment('center');
      return;
    }

    // Sort parties by type then name
    parties.sort((a, b) => {
      if (a.type !== b.type) return a.type.localeCompare(b.type);
      return a.name.localeCompare(b.name);
    });

    // Build rows with hyperlinks
    const rows = parties.map(party => {
      const sheetName = sanitizeSheetName(party.id);
      const link = '=HYPERLINK("#gid=' + getSheetGid(ss, sheetName) + '", "Open")';

      return [
        party.id,
        party.name,
        party.type,
        party.totalDebit || 0,
        party.totalCredit || 0,
        (party.totalDebit || 0) - (party.totalCredit || 0),
        party.lastTransaction || '',
        link
      ];
    });

    // Write data
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);

    // Format currency columns
    sheet.getRange(2, 4, rows.length, 3).setNumberFormat('₹ #,##0.00');

    // Format date column
    sheet.getRange(2, 7, rows.length, 1).setNumberFormat('dd-mm-yyyy');

    // Alternate row colors
    for (let i = 0; i < rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(2 + i, 1, 1, 8).setBackground('#f8f9fa');
      }
    }
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
    sanitizeSheetName
  };

})();
