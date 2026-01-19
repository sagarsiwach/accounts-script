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

    // Section 1: Organization Settings (using >> instead of = to avoid formula error)
    const orgSettings = [
      ['>> ORGANIZATION SETTINGS', '', ''],
      ['ORG_CODE', '', 'CM, KM, DEV, CPI, SM'],
      ['ORG_NAME', '', 'Full organization name'],
      ['FINANCIAL_YEAR', '2025-26', 'FY for this ledger'],
      ['CONTACTS_SHEET_ID', '1bqjiSyUUdfzV6AXbS13NiqyMlMUzKliXLKcn6zSdPCk', 'Contact Master Sheet ID'],
      ['CONTACTS_SHEET_NAME', 'ALL CONTACTS', 'Tab name for contacts'],
      ['', '', ''],
    ];

    // Section 2: Source Sheet IDs
    const sourceSettings = [
      ['>> SOURCE SHEETS', '', ''],
      ['PURCHASE_SHEET_ID', '', 'Google Sheet ID for Purchase Register'],
      ['PURCHASE_SHEET_NAME', 'Purchase Register', 'Tab name in the sheet'],
      ['SALES_SHEET_ID', '', 'Google Sheet ID for Sales Register'],
      ['SALES_SHEET_NAME', 'Sales Register', 'Tab name in the sheet'],
      ['BANK_SHEET_ID', '', 'Google Sheet ID for Bank workbook (all tabs scanned)'],
      ['', '', ''],
    ];

    // Section 3: Column Mappings - Purchase (prefilled with actual schema)
    const purchaseMapping = [
      ['>> PURCHASE COLUMN MAPPING', '', ''],
      ['PUR_LF_COL', 'L/F', 'Ledger folio / Party ID'],
      ['PUR_SUPPLIER_REF_COL', 'SUPPLIER REF', 'Supplier reference'],
      ['PUR_INWARD_NO_COL', 'INWARD NO.', 'Inward number'],
      ['PUR_INWARD_DATE_COL', 'INW DATE', 'Inward date'],
      ['PUR_INVOICE_NO_COL', 'INVOICE NO.', 'Invoice number'],
      ['PUR_INVOICE_DATE_COL', 'INVOICE DATE', 'Invoice date'],
      ['PUR_ACCOUNT_COL', 'ACCOUNT', 'Account'],
      ['PUR_SUPPLIER_NAME_COL', 'SUPPLIER NAME', 'Supplier name'],
      ['PUR_GST_NUMBER_COL', 'GST NUMBER', 'GST number'],
      ['PUR_HSN_CODE_COL', 'HSN CODE', 'HSN code'],
      ['PUR_ASS_VALUE_COL', 'ASS. VALUE', 'Assessable value'],
      ['PUR_IGST_COL', 'IGST', 'IGST amount'],
      ['PUR_CGST_COL', 'CGST', 'CGST amount'],
      ['PUR_SGST_COL', 'SGST', 'SGST amount'],
      ['PUR_OTHER_CHARGES_COL', 'OTHER CHARGES', 'Other charges'],
      ['PUR_GST_TOTAL_COL', 'GST TOTAL', 'GST total'],
      ['PUR_GRAND_TOTAL_COL', 'GRAND TOTAL', 'Grand total'],
      ['PUR_PURCHASE_YTD_COL', 'PURCHASE (YTD)', 'Purchase year to date'],
      ['PUR_GST_YTD_COL', 'GST (YTD)', 'GST year to date'],
      ['PUR_INVOICE_LINK_COL', 'INVOICE', 'Invoice file link'],
      ['PUR_GRN_LINK_COL', 'GRN', 'GRN file link'],
      ['PUR_PO_LINK_COL', 'PO', 'PO file link'],
      ['PUR_GST_TYPE_COL', 'GST TYPE', 'GST type'],
      ['PUR_TIN_COL', 'TIN', 'TIN'],
      ['PUR_GIN_COL', 'GIN', 'GIN'],
      ['PUR_GRN_NO_COL', 'GRN', 'GRN number'],
      ['PUR_SIN_COL', 'SIN', 'SIN'],
      ['PUR_REMARKS_COL', 'REMARKS', 'Remarks'],
      ['PUR_GST_REMARKS_COL', 'GST REMARKS', 'GST remarks'],
      ['PUR_GST_MONTH_COL', 'GST MONTH', 'GST month'],
      ['PUR_EXPENSE_COL', 'EXPENSE?', 'Is expense flag'],
      ['', '', ''],
    ];

    // Section 4: Column Mappings - Sales (prefilled with actual schema)
    const salesMapping = [
      ['>> SALES COLUMN MAPPING', '', ''],
      ['SAL_LF_COL', 'L/F', 'Ledger folio / Party ID'],
      ['SAL_CUSTOMER_NAME_COL', 'CUSTOMER NAME', 'Customer name'],
      ['SAL_INVOICE_NO_COL', 'INVOICE NO.', 'Invoice number'],
      ['SAL_INVOICE_DATE_COL', 'INVOICE DATE', 'Invoice date'],
      ['SAL_TYPE_COL', 'TYPE', 'Sale type'],
      ['SAL_BILL_TO_COL', 'BILL TO', 'Bill to'],
      ['SAL_BILLING_GST_COL', 'BILLING GST', 'Billing GST'],
      ['SAL_EWAY_BILL_COL', 'E-WAY BILL NUMBER', 'E-way bill number'],
      ['SAL_ASS_VALUE_COL', 'ASS. VALUE', 'Assessable value'],
      ['SAL_IGST_COL', 'IGST', 'IGST amount'],
      ['SAL_CGST_COL', 'CGST', 'CGST amount'],
      ['SAL_SGST_COL', 'SGST', 'SGST amount'],
      ['SAL_OTHER_CHARGES_COL', 'OTHER CHARGES', 'Other charges'],
      ['SAL_GST_TOTAL_COL', 'GST TOTAL', 'GST total'],
      ['SAL_GRAND_TOTAL_COL', 'GRAND TOTAL', 'Grand total'],
      ['SAL_SALE_YTD_COL', 'P. SALE (YEARLY)', 'Previous sale yearly'],
      ['SAL_GST_YTD_COL', 'P. GST (YEARLY)', 'Previous GST yearly'],
      ['SAL_VEH_NO_COL', 'VEH. NO', 'Vehicle number'],
      ['SAL_LR_NO_COL', 'LR NO.', 'LR number'],
      ['SAL_SRQ_COL', 'SRQ', 'SRQ'],
      ['SAL_INVOICE_LINK_COL', 'INVOICE', 'Invoice file link'],
      ['SAL_LR_COPY_LINK_COL', 'LR COPY', 'LR copy link'],
      ['SAL_GST_TYPE_COL', 'GST TYPE', 'GST type'],
      ['SAL_GST_MONTH_COL', 'GST MONTH', 'GST month'],
      ['SAL_REMARKS_COL', 'REMARKS', 'Remarks'],
      ['SAL_GST_REMARKS_COL', 'GST REMARKS', 'GST remarks'],
      ['', '', ''],
    ];

    // Section 5: Column Mappings - Bank (prefilled with actual schema)
    const bankMapping = [
      ['>> BANK COLUMN MAPPING', '', ''],
      ['BANK_DATE_COL', 'DATE', 'Transaction date'],
      ['BANK_PARTICULARS_COL', 'PARTICULARS', 'Particulars'],
      ['BANK_DEBIT_COL', 'DEBIT', 'Debit amount'],
      ['BANK_CREDIT_COL', 'CREDIT', 'Credit amount'],
      ['BANK_BALANCE_COL', 'BALANCE', 'Running balance'],
      ['BANK_LF_COL', 'L/F', 'Ledger folio / Party ID'],
      ['BANK_PARTY_NAME_COL', 'PARTY NAME', 'Party name'],
      ['BANK_PARTY_TYPE_COL', 'PARTY TYPE', 'Party type'],
      ['BANK_VOUCHER_TYPE_COL', 'VOUCHER TYPE', 'Voucher type'],
      ['BANK_REFERENCE_COL', 'REFERENCE', 'Reference'],
      ['BANK_ACCOUNT_ID_COL', 'ACCOUNT_ID', 'Account ID'],
      ['BANK_RECONCILED_COL', 'RECONCILED', 'Reconciliation status'],
      ['BANK_REMARKS_COL', 'REMARKS', 'Remarks'],
      ['', '', ''],
    ];

    // Section 6: Contact Master Schema Reference
    const contactSchema = [
      ['>> CONTACT MASTER SCHEMA', '', ''],
      ['CONTACT_SL_COL', 'SL', 'Serial / Contact ID (CG-SUP-0001, CG-CUS-0001, etc.)'],
      ['CONTACT_TYPE_COL', 'CONTACT TYPE', 'SUPPLIER, CUSTOMER, MASTER, DEALER, CONTRACTOR, RENTAL'],
      ['CONTACT_COMPANY_COL', 'COMPANY NAME', 'Company name'],
      ['CONTACT_ADDR1_COL', 'ADDRESS LINE 1', 'Address line 1'],
      ['CONTACT_ADDR2_COL', 'ADDRESS LINE 2', 'Address line 2'],
      ['CONTACT_DISTRICT_COL', 'DISTRICT', 'District'],
      ['CONTACT_STATE_COL', 'STATE', 'State'],
      ['CONTACT_PIN_COL', 'PIN CODE', 'PIN code'],
      ['CONTACT_GST_COL', 'GST', 'GST number'],
      ['CONTACT_RELATED_COL', 'RELATED COMPANY', 'Related / Parent company'],
      ['CONTACT_PERSON_COL', 'CONTACT PERSON', 'Contact person name'],
      ['CONTACT_MOBILE_COL', 'MOBILE NO', 'Mobile number'],
      ['CONTACT_EMAIL_COL', 'EMAIL ID', 'Email ID'],
      ['', '', ''],
    ];

    // Section 7: Notification Settings
    const notificationSettings = [
      ['>> NOTIFICATIONS', '', ''],
      ['ERROR_EMAIL', '', 'Email for error notifications'],
      ['LOG_RETENTION_DAYS', '30', 'Days to keep run logs'],
    ];

    // Combine all sections
    const allData = [
      ['SETTING', 'VALUE', 'DESCRIPTION'],
      ...orgSettings,
      ...sourceSettings,
      ...purchaseMapping,
      ...salesMapping,
      ...bankMapping,
      ...contactSchema,
      ...notificationSettings
    ];

    // Write to sheet
    const range = sheet.getRange(1, 1, allData.length, 3);
    range.setValues(allData);

    // Apply Roboto Condensed font to entire sheet
    sheet.getDataRange().setFontFamily('Roboto Condensed').setFontSize(8);

    // Set all rows to 20px height
    for (let i = 1; i <= allData.length; i++) {
      sheet.setRowHeight(i, 20);
    }

    // Format header row
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setFontSize(9);
    sheet.setRowHeight(1, 24);

    // Format section headers (using >> prefix now)
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).startsWith('>>')) {
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

      if (key && !key.startsWith('>>') && !key.startsWith('===')) {
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

    // Bank schema headers to detect (row 5)
    const BANK_SCHEMA_HEADERS = ['DATE', 'PARTICULARS', 'DEBIT', 'CREDIT', 'BALANCE', 'L/F'];

    try {
      const config = readConfig();

      // Test Purchase
      if (config.PURCHASE_SHEET_ID) {
        try {
          const sourceSheet = SpreadsheetApp.openById(config.PURCHASE_SHEET_ID);
          const tab = sourceSheet.getSheetByName(config.PURCHASE_SHEET_NAME || 'Purchase Register');
          if (tab) {
            const rowCount = tab.getLastRow();
            results.push({ name: 'Purchase', status: 'OK', message: rowCount + ' rows' });
          } else {
            results.push({ name: 'Purchase', status: 'ERROR', message: 'Tab "' + (config.PURCHASE_SHEET_NAME || 'Purchase Register') + '" not found' });
          }
        } catch (e) {
          results.push({ name: 'Purchase', status: 'ERROR', message: e.message });
        }
      } else {
        results.push({ name: 'Purchase', status: 'SKIPPED', message: 'Not configured' });
      }

      // Test Sales
      if (config.SALES_SHEET_ID) {
        try {
          const sourceSheet = SpreadsheetApp.openById(config.SALES_SHEET_ID);
          const tab = sourceSheet.getSheetByName(config.SALES_SHEET_NAME || 'Sales Register');
          if (tab) {
            const rowCount = tab.getLastRow();
            results.push({ name: 'Sales', status: 'OK', message: rowCount + ' rows' });
          } else {
            results.push({ name: 'Sales', status: 'ERROR', message: 'Tab "' + (config.SALES_SHEET_NAME || 'Sales Register') + '" not found' });
          }
        } catch (e) {
          results.push({ name: 'Sales', status: 'ERROR', message: e.message });
        }
      } else {
        results.push({ name: 'Sales', status: 'SKIPPED', message: 'Not configured' });
      }

      // Test Bank - scan ALL tabs for bank schema in row 5
      if (config.BANK_SHEET_ID) {
        try {
          const bankWorkbook = SpreadsheetApp.openById(config.BANK_SHEET_ID);
          const allSheets = bankWorkbook.getSheets();
          const bankTabs = [];
          let totalRows = 0;

          for (const sheet of allSheets) {
            try {
              // Check row 5 for bank schema headers
              const row5 = sheet.getRange(5, 1, 1, 6).getValues()[0];
              const headers = row5.map(h => String(h).trim().toUpperCase());

              // Check if first 6 columns match bank schema
              const matches = BANK_SCHEMA_HEADERS.every((h, i) => headers[i] === h);

              if (matches) {
                const rowCount = sheet.getLastRow() - 5; // Rows after header
                bankTabs.push(sheet.getName());
                totalRows += rowCount;
              }
            } catch (e) {
              // Skip sheets that can't be read
            }
          }

          if (bankTabs.length > 0) {
            results.push({
              name: 'Bank',
              status: 'OK',
              message: bankTabs.length + ' bank tabs found: ' + bankTabs.join(', ') + ' (' + totalRows + ' rows total)'
            });
          } else {
            results.push({
              name: 'Bank',
              status: 'ERROR',
              message: 'No tabs with bank schema in row 5. Expected: DATE, PARTICULARS, DEBIT, CREDIT, BALANCE, L/F'
            });
          }
        } catch (e) {
          results.push({ name: 'Bank', status: 'ERROR', message: e.message });
        }
      } else {
        results.push({ name: 'Bank', status: 'SKIPPED', message: 'Not configured' });
      }

      // Test Contacts
      if (config.CONTACTS_SHEET_ID) {
        try {
          const contactsSheet = SpreadsheetApp.openById(config.CONTACTS_SHEET_ID);
          const tab = contactsSheet.getSheetByName(config.CONTACTS_SHEET_NAME || 'ALL CONTACTS');
          if (tab) {
            const rowCount = tab.getLastRow();
            results.push({ name: 'Contacts', status: 'OK', message: rowCount + ' contacts' });
          } else {
            results.push({ name: 'Contacts', status: 'ERROR', message: 'Tab "' + (config.CONTACTS_SHEET_NAME || 'ALL CONTACTS') + '" not found' });
          }
        } catch (e) {
          results.push({ name: 'Contacts', status: 'ERROR', message: e.message });
        }
      } else {
        results.push({ name: 'Contacts', status: 'SKIPPED', message: 'Not configured' });
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
   * Gets all bank tabs from the configured bank sheet
   * Returns array of {name, sheet} objects for tabs matching bank schema
   */
  function getBankTabs() {
    const config = readConfig();
    const BANK_SCHEMA_HEADERS = ['DATE', 'PARTICULARS', 'DEBIT', 'CREDIT', 'BALANCE', 'L/F'];
    const bankTabs = [];

    if (!config.BANK_SHEET_ID) return bankTabs;

    try {
      const bankWorkbook = SpreadsheetApp.openById(config.BANK_SHEET_ID);
      const allSheets = bankWorkbook.getSheets();

      for (const sheet of allSheets) {
        try {
          const row5 = sheet.getRange(5, 1, 1, 6).getValues()[0];
          const headers = row5.map(h => String(h).trim().toUpperCase());
          const matches = BANK_SCHEMA_HEADERS.every((h, i) => headers[i] === h);

          if (matches) {
            bankTabs.push({ name: sheet.getName(), sheet: sheet });
          }
        } catch (e) {
          // Skip
        }
      }
    } catch (e) {
      Logger.log('Error getting bank tabs: ' + e.message);
    }

    return bankTabs;
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
    getBankTabs,
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
