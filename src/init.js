/**
 * Initialization Module
 * Creates all required tabs and structures for a new org ledger sheet
 *
 * @fileoverview INIT functionality for CG Accounts
 */

const Init = (function() {

  /**
   * Main initialization function - creates all required tabs
   * Run this on a fresh sheet to set up the structure
   */
  function initialize() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    try {
      // Create all required tabs
      createConfigTab(ss);
      createLedgerMasterTab(ss);
      createRunLogTab(ss);

      // Test connections (will show errors if sources not configured)
      const configValid = validateConfig();

      if (configValid.valid) {
        ui.alert('Initialization Complete',
          'All tabs created successfully.\n\n' +
          'Next steps:\n' +
          '1. Fill in CONFIG tab with source sheet IDs\n' +
          '2. Run "Test Connection" to verify\n' +
          '3. Run "Refresh" to fetch data',
          ui.ButtonSet.OK);
      } else {
        ui.alert('Initialization Complete',
          'Tabs created. Please configure the CONFIG tab:\n\n' +
          '• Set ORG_CODE\n' +
          '• Add source sheet IDs\n' +
          '• Configure column mappings\n\n' +
          'Then run "Test Connection".',
          ui.ButtonSet.OK);
      }

    } catch (error) {
      ui.alert('Initialization Failed', error.message, ui.ButtonSet.OK);
      Logger.log('Init error: ' + error.message);
    }
  }

  /**
   * Creates the CONFIG tab with all settings
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  function createConfigTab(ss) {
    let sheet = ss.getSheetByName('CONFIG');

    if (!sheet) {
      sheet = ss.insertSheet('CONFIG');
    } else {
      sheet.clear();
    }

    // Move to first position
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);

    // Section 1: Organization Settings
    const orgSettings = [
      ['=== ORGANIZATION SETTINGS ===', '', ''],
      ['ORG_CODE', '', 'CM, KM, DEV, CPI, SM'],
      ['ORG_NAME', '', 'Full organization name'],
      ['FINANCIAL_YEAR', '2025-26', 'FY for this ledger'],
      ['', '', ''],
    ];

    // Section 2: Source Sheet IDs
    const sourceSettings = [
      ['=== SOURCE SHEETS ===', '', ''],
      ['PURCHASE_SHEET_ID', '', 'Google Sheet ID for Purchase Register'],
      ['PURCHASE_SHEET_NAME', 'Purchase Register', 'Tab name in the sheet'],
      ['SALES_SHEET_ID', '', 'Google Sheet ID for Sales Register'],
      ['SALES_SHEET_NAME', 'Sales Register', 'Tab name in the sheet'],
      ['BANK_SHEET_ID', '', 'Google Sheet ID for Bank Statement'],
      ['BANK_SHEET_NAME', 'Bank Statement', 'Tab name in the sheet'],
      ['SUPPLIER_MASTER_ID', '', 'Google Sheet ID for Supplier Master'],
      ['SUPPLIER_MASTER_NAME', 'Supplier Master', 'Tab name'],
      ['CUSTOMER_MASTER_ID', '', 'Google Sheet ID for Customer Master'],
      ['CUSTOMER_MASTER_NAME', 'Customer Master', 'Tab name'],
      ['', '', ''],
    ];

    // Section 3: Column Mappings - Purchase
    const purchaseMapping = [
      ['=== PURCHASE COLUMN MAPPING ===', '', ''],
      ['PUR_DATE_COL', 'INW_DATE', 'Date column'],
      ['PUR_PARTY_ID_COL', 'L/F', 'Party ID column'],
      ['PUR_PARTY_NAME_COL', 'SUPPLIER_NAME', 'Party name column'],
      ['PUR_INVOICE_COL', 'INVOICE_NO', 'Invoice number column'],
      ['PUR_AMOUNT_COL', 'GRAND_TOTAL', 'Total amount column'],
      ['PUR_GST_COL', 'GST_TOTAL', 'GST total column'],
      ['PUR_IGST_COL', 'IGST', 'IGST column'],
      ['PUR_CGST_COL', 'CGST', 'CGST column'],
      ['PUR_SGST_COL', 'SGST', 'SGST column'],
      ['PUR_REMARKS_COL', 'REMARKS', 'Remarks column'],
      ['', '', ''],
    ];

    // Section 4: Column Mappings - Sales
    const salesMapping = [
      ['=== SALES COLUMN MAPPING ===', '', ''],
      ['SAL_DATE_COL', 'INVOICE_DATE', 'Date column'],
      ['SAL_PARTY_ID_COL', 'L/F', 'Party ID column'],
      ['SAL_PARTY_NAME_COL', 'PARTY_NAME', 'Party name column'],
      ['SAL_INVOICE_COL', 'INVOICE_NO', 'Invoice number column'],
      ['SAL_AMOUNT_COL', 'GRAND_TOTAL', 'Total amount column'],
      ['SAL_GST_COL', 'GST_TOTAL', 'GST total column'],
      ['SAL_REMARKS_COL', 'REMARKS', 'Remarks column'],
      ['', '', ''],
    ];

    // Section 5: Column Mappings - Bank
    const bankMapping = [
      ['=== BANK COLUMN MAPPING ===', '', ''],
      ['BANK_DATE_COL', 'DATE', 'Date column'],
      ['BANK_PARTY_ID_COL', 'L/F', 'Party ID column'],
      ['BANK_PARTY_NAME_COL', 'PARTY NAME', 'Party name column'],
      ['BANK_DEBIT_COL', 'DEBIT', 'Debit column'],
      ['BANK_CREDIT_COL', 'CREDIT', 'Credit column'],
      ['BANK_PARTICULARS_COL', 'PARTICULARS', 'Particulars column'],
      ['BANK_REF_COL', 'REFERENCE', 'Reference column'],
      ['BANK_VOUCHER_COL', 'VOUCHER TYPE', 'Voucher type column'],
      ['', '', ''],
    ];

    // Section 6: Notification Settings
    const notificationSettings = [
      ['=== NOTIFICATIONS ===', '', ''],
      ['ERROR_EMAIL', '', 'Email for error notifications'],
      ['LOG_RETENTION_DAYS', '30', 'Days to keep run logs'],
      ['', '', ''],
    ];

    // Section 7: Party Master Fields (for ledger header)
    const partyFields = [
      ['=== PARTY MASTER FIELDS ===', '', ''],
      ['PARTY_PARENT_COL', 'PARENT_COMPANY', 'Parent company column'],
      ['PARTY_ADDRESS1_COL', 'ADDRESS_LINE_1', 'Address line 1'],
      ['PARTY_ADDRESS2_COL', 'ADDRESS_LINE_2', 'Address line 2'],
      ['PARTY_CONTACT_COL', 'CONTACT', 'Contact info column'],
      ['PARTY_GST_COL', 'GST_NUMBER', 'GST number column'],
      ['PARTY_EMAIL_COL', 'EMAIL', 'Email column'],
    ];

    // Combine all sections
    const allData = [
      ['SETTING', 'VALUE', 'DESCRIPTION'],
      ...orgSettings,
      ...sourceSettings,
      ...purchaseMapping,
      ...salesMapping,
      ...bankMapping,
      ...notificationSettings,
      ...partyFields
    ];

    // Write to sheet
    const range = sheet.getRange(1, 1, allData.length, 3);
    range.setValues(allData);

    // Format header
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white');

    // Format section headers
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).startsWith('===')) {
        sheet.getRange(i + 1, 1, 1, 3)
          .setFontWeight('bold')
          .setBackground('#d9ead3')
          .setFontColor('#274e13');
      }
    }

    // Set column widths
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 250);

    // Freeze header
    sheet.setFrozenRows(1);
  }

  /**
   * Creates the Ledger Master tab (index of all ledgers)
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  function createLedgerMasterTab(ss) {
    let sheet = ss.getSheetByName('Ledger Master');

    if (!sheet) {
      sheet = ss.insertSheet('Ledger Master');
    } else {
      sheet.clear();
    }

    // Move to second position
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(2);

    // Headers
    const headers = [
      'PARTY CODE', 'PARTY NAME', 'TYPE', 'TOTAL DEBIT', 'TOTAL CREDIT',
      'BALANCE', 'LAST TRANSACTION', 'LINK'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setHorizontalAlignment('center');

    // Set column widths
    sheet.setColumnWidth(1, 120); // Party Code
    sheet.setColumnWidth(2, 250); // Party Name
    sheet.setColumnWidth(3, 100); // Type
    sheet.setColumnWidth(4, 120); // Total Debit
    sheet.setColumnWidth(5, 120); // Total Credit
    sheet.setColumnWidth(6, 120); // Balance
    sheet.setColumnWidth(7, 130); // Last Transaction
    sheet.setColumnWidth(8, 80);  // Link

    // Freeze header
    sheet.setFrozenRows(1);

    // Add placeholder message
    sheet.getRange(3, 1, 1, headers.length)
      .merge()
      .setValue('Run "Refresh" to populate ledger data')
      .setFontStyle('italic')
      .setFontColor('#666666')
      .setHorizontalAlignment('center');
  }

  /**
   * Creates the RUN_LOG tab
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  function createRunLogTab(ss) {
    let sheet = ss.getSheetByName('RUN_LOG');

    if (!sheet) {
      sheet = ss.insertSheet('RUN_LOG');
    } else {
      sheet.clear();
    }

    // Headers
    const headers = [
      'TIMESTAMP', 'SOURCE', 'ACTION', 'ROWS_FETCHED', 'ROWS_WRITTEN',
      'STATUS', 'DURATION_MS', 'ERROR_MESSAGE'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white');

    // Set column widths
    sheet.setColumnWidth(1, 150); // Timestamp
    sheet.setColumnWidth(2, 100); // Source
    sheet.setColumnWidth(3, 100); // Action
    sheet.setColumnWidth(4, 100); // Rows Fetched
    sheet.setColumnWidth(5, 100); // Rows Written
    sheet.setColumnWidth(6, 80);  // Status
    sheet.setColumnWidth(7, 100); // Duration
    sheet.setColumnWidth(8, 300); // Error

    // Freeze header
    sheet.setFrozenRows(1);
  }

  /**
   * Reads configuration from CONFIG tab
   * @returns {Object} Configuration object
   */
  function readConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CONFIG');

    if (!sheet) {
      throw new Error('CONFIG tab not found. Run Initialize first.');
    }

    const data = sheet.getDataRange().getValues();
    const config = {};

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      const value = data[i][1];

      if (key && !key.startsWith('===')) {
        config[key] = value;
      }
    }

    return config;
  }

  /**
   * Validates configuration completeness
   * @returns {Object} Validation result
   */
  function validateConfig() {
    const errors = [];
    const warnings = [];

    try {
      const config = readConfig();

      // Required fields
      if (!config.ORG_CODE) errors.push('ORG_CODE not set');
      if (!config.PURCHASE_SHEET_ID && !config.SALES_SHEET_ID && !config.BANK_SHEET_ID) {
        errors.push('At least one source sheet ID required');
      }

    } catch (e) {
      errors.push(e.message);
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Tests connection to all configured source sheets
   */
  function testConnection() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const results = [];

    try {
      const config = readConfig();

      // Test each source
      const sources = [
        { name: 'Purchase', idKey: 'PURCHASE_SHEET_ID', nameKey: 'PURCHASE_SHEET_NAME' },
        { name: 'Sales', idKey: 'SALES_SHEET_ID', nameKey: 'SALES_SHEET_NAME' },
        { name: 'Bank', idKey: 'BANK_SHEET_ID', nameKey: 'BANK_SHEET_NAME' },
        { name: 'Supplier Master', idKey: 'SUPPLIER_MASTER_ID', nameKey: 'SUPPLIER_MASTER_NAME' },
        { name: 'Customer Master', idKey: 'CUSTOMER_MASTER_ID', nameKey: 'CUSTOMER_MASTER_NAME' }
      ];

      for (const source of sources) {
        const sheetId = config[source.idKey];
        const sheetName = config[source.nameKey];

        if (!sheetId) {
          results.push({ name: source.name, status: 'SKIPPED', message: 'Not configured' });
          continue;
        }

        try {
          const sourceSheet = SpreadsheetApp.openById(sheetId);
          const tab = sourceSheet.getSheetByName(sheetName);

          if (tab) {
            const rowCount = tab.getLastRow();
            results.push({ name: source.name, status: 'OK', message: rowCount + ' rows' });
          } else {
            results.push({ name: source.name, status: 'ERROR', message: 'Tab "' + sheetName + '" not found' });
          }
        } catch (e) {
          results.push({ name: source.name, status: 'ERROR', message: e.message });
        }
      }

      // Show results
      let message = 'Connection Test Results:\n\n';
      for (const r of results) {
        const icon = r.status === 'OK' ? '✓' : (r.status === 'ERROR' ? '✗' : '○');
        message += icon + ' ' + r.name + ': ' + r.status + '\n   ' + r.message + '\n\n';
      }

      ui.alert('Connection Test', message, ui.ButtonSet.OK);

      // Log results
      logRun('TEST', 'Connection Test', results.filter(r => r.status === 'OK').length,
        0, results.every(r => r.status !== 'ERROR') ? 'SUCCESS' : 'PARTIAL', 0);

    } catch (error) {
      ui.alert('Test Failed', error.message, ui.ButtonSet.OK);
    }
  }

  /**
   * Writes a log entry to RUN_LOG
   */
  function logRun(source, action, rowsFetched, rowsWritten, status, durationMs, errorMessage) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('RUN_LOG');

      if (!sheet) return;

      sheet.appendRow([
        new Date(),
        source,
        action,
        rowsFetched || 0,
        rowsWritten || 0,
        status,
        durationMs || 0,
        errorMessage || ''
      ]);

    } catch (e) {
      Logger.log('Failed to write log: ' + e.message);
    }
  }

  // Public API
  return {
    initialize,
    readConfig,
    validateConfig,
    testConnection,
    logRun,
    createConfigTab,
    createLedgerMasterTab,
    createRunLogTab
  };

})();

// Global function for menu
function initializeSheet() {
  Init.initialize();
}

function testConnection() {
  Init.testConnection();
}
