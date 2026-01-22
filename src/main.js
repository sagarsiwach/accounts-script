/**
 * Classic Group Accounts Script
 * Main entry point - menu and trigger handlers
 *
 * @fileoverview Entry points for the accounts aggregation system
 * @author Sagar Siwach
 * @version 0.6.0
 *
 * Architecture: Each org's Ledger sheet is self-contained with:
 * - CONFIG tab: All settings (sources, mappings)
 * - Ledger Master tab: Index of all party ledgers with hyperlinks
 * - RUN_LOG tab: Execution history
 * - Individual party ledger tabs (CG-SUP-0001, CG-CUS-0001, etc.)
 *
 * Data Sources:
 * - Purchase Register: Supplier purchase transactions
 * - Sales Register: Customer sales transactions
 * - Bank Statement: Bank transactions (multi-tab support, schema in row 5)
 * - Contacts: Central contacts sheet for party master data
 */

/**
 * Parses a numeric value, handling Indian number format with commas
 * @param {*} val - Value to parse (number or string)
 * @returns {number} Parsed number or 0
 */
function parseNumericValue(val) {
  if (!val) return 0;
  if (typeof val === 'number') return val;
  // Remove currency symbols, commas, and spaces
  const cleaned = String(val).replace(/[₹$,\s]/g, '');
  return parseFloat(cleaned) || 0;
}

/**
 * Runs when the add-on is installed
 * Required for add-ons to show menu immediately after installation
 * @param {Object} e - Event object
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Creates the custom menu when spreadsheet opens
 * Works for both container-bound and add-on installations
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createAddonMenu()
    .addItem('Open Dashboard', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Initialize Sheet', 'initializeSheet')
      .addItem('Test Connection', 'testConnection'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Ledgers')
      .addItem('Create All Ledgers', 'createAllLedgers')
      .addItem('Create Single Ledger...', 'createSingleLedgerPrompt'))
    .addSeparator()
    .addItem('Refresh Data', 'refreshData')
    .addSeparator()
    .addSubMenu(ui.createMenu('View')
      .addItem('Run Logs', 'showLogs')
      .addItem('Config', 'goToConfig'));

  menu.addToUi();
}

/**
 * Add-on homepage trigger - shows card UI in sidebar
 * @param {Object} e - Event object
 * @returns {Card} Card to display
 */
function onHomepage(e) {
  return createHomepageCard();
}

/**
 * Creates the homepage card for the add-on
 * @returns {Card} The card to display
 */
function createHomepageCard() {
  const builder = CardService.newCardBuilder();

  // Header
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle('CG Accounts')
      .setSubtitle('Financial Data Aggregation')
      .setImageStyle(CardService.ImageStyle.SQUARE)
  );

  // Check if sheet is initialized
  let configStatus = 'Not initialized';
  let orgCode = '-';
  try {
    const config = Init.readConfig();
    if (config && config.ORG_CODE) {
      configStatus = 'Configured';
      orgCode = config.ORG_CODE;
    }
  } catch (e) {
    configStatus = 'Not initialized';
  }

  // Status Section
  const statusSection = CardService.newCardSection()
    .setHeader('Status')
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel('Organization')
        .setText(orgCode)
    )
    .addWidget(
      CardService.newDecoratedText()
        .setTopLabel('Config Status')
        .setText(configStatus)
    );

  builder.addSection(statusSection);

  // Actions Section
  const actionsSection = CardService.newCardSection()
    .setHeader('Quick Actions')
    .addWidget(
      CardService.newTextButton()
        .setText('Initialize Sheet')
        .setOnClickAction(
          CardService.newAction().setFunctionName('initializeSheet')
        )
    )
    .addWidget(
      CardService.newTextButton()
        .setText('Refresh Data')
        .setOnClickAction(
          CardService.newAction().setFunctionName('refreshData')
        )
    )
    .addWidget(
      CardService.newTextButton()
        .setText('Test Connection')
        .setOnClickAction(
          CardService.newAction().setFunctionName('testConnection')
        )
    )
    .addWidget(
      CardService.newTextButton()
        .setText('Open Full Dashboard')
        .setOnClickAction(
          CardService.newAction().setFunctionName('showSidebar')
        )
    );

  builder.addSection(actionsSection);

  return builder.build();
}

/**
 * Installable trigger for hourly refresh
 */
function hourlyRefresh() {
  const startTime = new Date();
  Logger.log('Starting hourly refresh at ' + startTime.toISOString());

  try {
    const result = refreshData();
    Logger.log('Hourly refresh completed: ' + JSON.stringify(result));
  } catch (error) {
    Logger.log('Hourly refresh failed: ' + error.message);
    sendErrorNotification('Hourly Refresh Failed', error);
  }
}

/**
 * Shows the main sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ui/Sidebar')
    .setTitle('CG Accounts')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the logs panel
 */
function showLogs() {
  const html = HtmlService.createHtmlOutputFromFile('ui/Logs')
    .setTitle('Run Logs')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Navigates to CONFIG tab
 */
function goToConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CONFIG');
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('CONFIG tab not found. Run Initialize first.');
  }
}

/**
 * Initialize the sheet structure
 */
function initializeSheet() {
  Init.initialize();
}

/**
 * Test connection to source sheets
 */
function testConnection() {
  Init.testConnection();
}

/**
 * Main refresh function - fetches all data and generates ledgers
 */
function refreshData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date();

  const result = {
    purchase: { rows: 0, status: 'pending' },
    sales: { rows: 0, status: 'pending' },
    bank: { rows: 0, status: 'pending' },
    ledgers: { count: 0, status: 'pending' }
  };

  try {
    // Read config from local CONFIG tab
    const config = Init.readConfig();

    if (!config.ORG_CODE) {
      ui.alert('Configuration Error', 'ORG_CODE not set in CONFIG tab.', ui.ButtonSet.OK);
      return result;
    }

    // Fetch from sources (sidebar handles its own loading UI)
    if (config.PURCHASE_SHEET_ID) {
      result.purchase = fetchSourceData(config, 'PURCHASE');
      Init.logRun('PURCHASE', 'Fetch', result.purchase.rows, 0,
        result.purchase.status, Date.now() - startTime, result.purchase.error);
    }

    if (config.SALES_SHEET_ID) {
      result.sales = fetchSourceData(config, 'SALES');
      Init.logRun('SALES', 'Fetch', result.sales.rows, 0,
        result.sales.status, Date.now() - startTime, result.sales.error);
    }

    if (config.BANK_SHEET_ID) {
      const bankTransactions = fetchBankDataAllTabs(config);
      result.bank.data = bankTransactions;
      result.bank.rows = bankTransactions.length;
      result.bank.status = 'SUCCESS';
      Init.logRun('BANK', 'Fetch', result.bank.rows, 0,
        result.bank.status, Date.now() - startTime, result.bank.error);
    }

    // Combine all data and generate ledgers
    const allTransactions = [
      ...(result.purchase.data || []),
      ...(result.sales.data || []),
      ...(result.bank.data || [])
    ];

    if (allTransactions.length > 0) {
      result.ledgers = generateAllLedgers(ss, config, allTransactions);
      Init.logRun('LEDGER', 'Generate', 0, result.ledgers.count,
        result.ledgers.status, Date.now() - startTime, result.ledgers.error);
    }

    // Show completion
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    ui.alert('Refresh Complete',
      'Org: ' + config.ORG_CODE + '\n\n' +
      'Purchase: ' + result.purchase.rows + ' rows\n' +
      'Sales: ' + result.sales.rows + ' rows\n' +
      'Bank: ' + result.bank.rows + ' rows\n\n' +
      'Ledgers generated: ' + result.ledgers.count + '\n' +
      'Duration: ' + duration + 's',
      ui.ButtonSet.OK);

  } catch (error) {
    Init.logRun('REFRESH', 'Error', 0, 0, 'ERROR', Date.now() - startTime, error.message);
    ui.alert('Refresh Failed', error.message, ui.ButtonSet.OK);
  }

  return result;
}

/**
 * Fetches data from a source sheet
 * @param {Object} config - Configuration object
 * @param {string} sourceType - PURCHASE, SALES, or BANK
 * @returns {Object} Result with data array
 */
function fetchSourceData(config, sourceType) {
  const result = { data: [], rows: 0, status: 'pending', error: null };

  try {
    // Get source sheet ID and name
    const sheetId = config[sourceType + '_SHEET_ID'];
    const sheetName = config[sourceType + '_SHEET_NAME'];

    if (!sheetId) {
      result.status = 'SKIPPED';
      result.error = 'Not configured';
      return result;
    }

    // Open source sheet
    const sourceSheet = SpreadsheetApp.openById(sheetId);
    const tab = sourceSheet.getSheetByName(sheetName);

    if (!tab) {
      throw new Error('Tab "' + sheetName + '" not found in source sheet');
    }

    // Get all data
    const allData = tab.getDataRange().getValues();
    if (allData.length < 2) {
      result.status = 'SUCCESS';
      result.rows = 0;
      return result;
    }

    // Auto-detect header row by looking for known columns
    const headerRow = detectHeaderRow(allData, sourceType);
    if (headerRow === -1) {
      throw new Error('Could not detect header row. Looking for "L/F" column.');
    }

    // Get column mappings for this source type
    const mapping = getColumnMapping(config, sourceType);

    // Headers are at the detected row (0-indexed)
    const headers = allData[headerRow];
    const rows = allData.slice(headerRow + 1);  // Data starts after header row

    // Transform each row
    result.data = rows
      .filter(row => row.some(cell => cell !== ''))
      .map(row => transformRow(row, headers, mapping, sourceType));

    result.rows = result.data.length;
    result.status = 'SUCCESS';

  } catch (error) {
    result.status = 'ERROR';
    result.error = error.message;
    Logger.log('Fetch error for ' + sourceType + ': ' + error.message);
  }

  return result;
}

/**
 * Gets column mapping from config for a source type
 * Maps CONFIG keys to standardized field names
 */
function getColumnMapping(config, sourceType) {
  if (sourceType === 'PURCHASE') {
    return {
      date: config.PUR_INVOICE_DATE_COL || config.PUR_INWARD_DATE_COL,
      partyId: config.PUR_LF_COL,  // L/F column contains party ID
      partyName: config.PUR_SUPPLIER_NAME_COL,
      invoice: config.PUR_INVOICE_NO_COL,
      amount: config.PUR_GRAND_TOTAL_COL,
      gst: config.PUR_GST_TOTAL_COL,
      particulars: config.PUR_ACCOUNT_COL,
      remarks: config.PUR_REMARKS_COL
    };
  } else if (sourceType === 'SALES') {
    return {
      date: config.SAL_INVOICE_DATE_COL,
      partyId: config.SAL_LF_COL,  // L/F column contains party ID
      partyName: config.SAL_CUSTOMER_NAME_COL,
      invoice: config.SAL_INVOICE_NO_COL,
      amount: config.SAL_GRAND_TOTAL_COL,
      gst: config.SAL_GST_TOTAL_COL,
      particulars: config.SAL_TYPE_COL,
      remarks: config.SAL_REMARKS_COL
    };
  } else {
    // BANK
    return {
      date: config.BANK_DATE_COL,
      partyId: config.BANK_LF_COL,
      partyName: config.BANK_PARTY_NAME_COL,
      debit: config.BANK_DEBIT_COL,
      credit: config.BANK_CREDIT_COL,
      particulars: config.BANK_PARTICULARS_COL,
      voucherType: config.BANK_VOUCHER_TYPE_COL,
      reference: config.BANK_REFERENCE_COL,
      remarks: config.BANK_REMARKS_COL
    };
  }
}

/**
 * Auto-detects which row contains the headers by scanning for known column names
 * @param {Array} allData - All data from the sheet
 * @param {string} sourceType - PURCHASE, SALES, or BANK
 * @returns {number} 0-indexed row number, or -1 if not found
 */
function detectHeaderRow(allData, sourceType) {
  // Define key columns to look for based on source type
  const keyColumns = {
    PURCHASE: ['L/F', 'SUPPLIER NAME', 'INVOICE NO', 'GRAND TOTAL'],
    SALES: ['L/F', 'CUSTOMER NAME', 'INVOICE NO', 'GRAND TOTAL'],
    BANK: ['DATE', 'PARTICULARS', 'DEBIT', 'CREDIT', 'BALANCE', 'L/F']
  };

  const columnsToFind = keyColumns[sourceType] || ['L/F'];

  // Scan first 10 rows to find header row
  const maxRowsToScan = Math.min(allData.length, 10);

  for (let rowIdx = 0; rowIdx < maxRowsToScan; rowIdx++) {
    const row = allData[rowIdx];
    const rowHeaders = row.map(cell => String(cell).trim().toUpperCase());

    // Check if this row contains at least 2 of the key columns (or L/F at minimum)
    let matchCount = 0;
    let hasLF = false;

    for (const col of columnsToFind) {
      const colUpper = col.toUpperCase();
      // Check for exact match or partial match (e.g., "INVOICE NO." matches "INVOICE NO")
      const found = rowHeaders.some(h =>
        h === colUpper ||
        h.replace(/[.\s]/g, '') === colUpper.replace(/[.\s]/g, '')
      );
      if (found) {
        matchCount++;
        if (colUpper === 'L/F') hasLF = true;
      }
    }

    // Accept row if it has L/F and at least 1 other key column, OR has 3+ matches
    if ((hasLF && matchCount >= 2) || matchCount >= 3) {
      Logger.log('Detected header row at row ' + (rowIdx + 1) + ' for ' + sourceType);
      return rowIdx;
    }
  }

  // Fallback: just look for L/F column anywhere in first 10 rows
  for (let rowIdx = 0; rowIdx < maxRowsToScan; rowIdx++) {
    const row = allData[rowIdx];
    const hasLF = row.some(cell =>
      String(cell).trim().toUpperCase() === 'L/F'
    );
    if (hasLF) {
      Logger.log('Detected header row at row ' + (rowIdx + 1) + ' (fallback L/F match) for ' + sourceType);
      return rowIdx;
    }
  }

  return -1;  // Not found
}

/**
 * Transforms a source row to standardized format
 */
function transformRow(row, headers, mapping, sourceType) {
  function getVal(columnName) {
    if (!columnName) return null;
    const idx = headers.findIndex(h =>
      String(h).trim().toUpperCase() === String(columnName).trim().toUpperCase()
    );
    return idx >= 0 ? row[idx] : null;
  }

  function parseNum(val) {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    const cleaned = String(val).replace(/[₹$,\s]/g, '');
    return parseFloat(cleaned) || 0;
  }

  const entry = {
    date: getVal(mapping.date),
    partyId: String(getVal(mapping.partyId) || ''),
    partyName: String(getVal(mapping.partyName) || ''),
    docNo: String(getVal(mapping.invoice) || getVal(mapping.reference) || ''),
    docType: sourceType,
    voucherType: String(getVal(mapping.voucherType) || sourceType),
    particulars: String(getVal(mapping.particulars) || getVal(mapping.remarks) || ''),
    debit: 0,
    credit: 0,
    reference: String(getVal(mapping.reference) || '')
  };

  // Handle amounts based on source type
  if (sourceType === 'PURCHASE') {
    entry.credit = parseNum(getVal(mapping.amount));
    entry.particulars = entry.particulars || 'Purchase Invoice: ' + entry.docNo;
  } else if (sourceType === 'SALES') {
    entry.debit = parseNum(getVal(mapping.amount));
    entry.particulars = entry.particulars || 'Sales Invoice: ' + entry.docNo;
  } else if (sourceType === 'BANK') {
    entry.debit = parseNum(getVal(mapping.debit));
    entry.credit = parseNum(getVal(mapping.credit));
  }

  return entry;
}

/**
 * Generates all party ledgers from transactions
 * Separates into [SU] Supplier and [CU] Customer ledgers
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Object} config - Configuration object
 * @param {Array} transactions - Array of all transactions
 * @returns {Object} Result with count, status, and error
 */
function generateAllLedgers(ss, config, transactions) {
  const result = { count: 0, status: 'pending', error: null };

  try {
    // Get company info for ledger headers
    const company = getCompanyFromConfig(config);

    // Batch load all contacts once (performance optimization)
    const allContacts = fetchAllContacts(config);

    // Group transactions by party AND ledger category
    // Key format: "partyId|category" where category is SU or CU
    const ledgerMap = {};

    for (const txn of transactions) {
      if (!txn.partyId) continue;

      const partyIdUpper = String(txn.partyId).toUpperCase();

      // Determine ledger category based on source and party type
      const ledgerCategory = determineLedgerCategory(partyIdUpper, txn.docType, txn);

      // Skip if category is null (e.g., contractors)
      if (!ledgerCategory) continue;

      const ledgerKey = partyIdUpper + '|' + ledgerCategory;

      if (!ledgerMap[ledgerKey]) {
        ledgerMap[ledgerKey] = {
          id: partyIdUpper,
          name: txn.partyName || '',
          ledgerCategory: ledgerCategory,
          type: ledgerCategory === 'SU' ? 'SUPPLIER' : 'CUSTOMER',
          transactions: [],
          totalDebit: 0,
          totalCredit: 0,
          lastTransaction: null
        };
      }

      ledgerMap[ledgerKey].transactions.push(txn);
      ledgerMap[ledgerKey].totalDebit += parseFloat(txn.debit) || 0;
      ledgerMap[ledgerKey].totalCredit += parseFloat(txn.credit) || 0;

      if (!ledgerMap[ledgerKey].lastTransaction ||
          txn.date > ledgerMap[ledgerKey].lastTransaction) {
        ledgerMap[ledgerKey].lastTransaction = txn.date;
      }
    }

    const ledgers = Object.values(ledgerMap);

    // Enrich party info from contacts (using pre-loaded data)
    for (const ledger of ledgers) {
      const contactInfo = allContacts[ledger.id];
      if (contactInfo) {
        ledger.name = contactInfo.name || ledger.name;
        ledger.address1 = contactInfo.address1 || '';
        ledger.address2 = contactInfo.address2 || '';
        ledger.gst = contactInfo.gst || '';
        ledger.phone = contactInfo.phone || '';
        ledger.email = contactInfo.email || '';
      }
    }

    // Create individual ledger sheets
    for (const ledger of ledgers) {
      // Sort transactions by date
      ledger.transactions.sort((a, b) => {
        const dateA = a.date instanceof Date ? a.date : new Date(a.date || 0);
        const dateB = b.date instanceof Date ? b.date : new Date(b.date || 0);
        return dateA - dateB;
      });

      PartyLedger.createPartyLedger(ss, ledger, company, ledger.transactions, ledger.ledgerCategory);
      result.count++;
    }

    // Update Ledger Master index
    PartyLedger.updateLedgerMasterIndex(ss, ledgers);

    // Reorder tabs: Ledger Master first, ledgers in middle, CONFIG and RUN_LOG at end
    Init.reorderTabs(ss);

    result.status = 'SUCCESS';

  } catch (error) {
    result.status = 'ERROR';
    result.error = error.message;
    Logger.log('Ledger generation error: ' + error.message);
  }

  return result;
}

/**
 * Determines the ledger category (SU or CU) based on party ID and transaction source
 * @param {string} partyId - Party ID (uppercase)
 * @param {string} docType - Document type: PURCHASE, SALES, or BANK
 * @param {Object} txn - Transaction object (for bank debit/credit check)
 * @returns {string|null} 'SU' for supplier, 'CU' for customer, null to skip
 */
function determineLedgerCategory(partyId, docType, txn) {
  // Source-based routing takes priority
  if (docType === 'PURCHASE') {
    return 'SU';  // Purchase = Supplier ledger
  }

  if (docType === 'SALES') {
    return 'CU';  // Sales = Customer ledger
  }

  // For BANK transactions, route based on party ID prefix
  if (docType === 'BANK') {
    if (partyId.includes('-SUP-')) {
      return 'SU';
    }
    if (partyId.includes('-CUS-') || partyId.includes('-REN-') || partyId.includes('-DEA-')) {
      return 'CU';
    }
    if (partyId.includes('-CON-')) {
      return null;  // Skip contractors for now
    }
    if (partyId.includes('-MAS-')) {
      // For Master parties in bank, use debit/credit to determine
      // Credit to bank (we paid them) = Supplier
      // Debit to bank (they paid us) = Customer
      if ((txn.credit || 0) > 0) {
        return 'SU';  // Payment made = Supplier
      }
      if ((txn.debit || 0) > 0) {
        return 'CU';  // Payment received = Customer
      }
    }
    // Default bank transactions without clear party type to Customer
    return 'CU';
  }

  // Default fallback
  return 'CU';
}

/**
 * Sends error notification email
 */
function sendErrorNotification(subject, error) {
  try {
    const config = Init.readConfig();
    const email = config.ERROR_EMAIL;

    if (!email) return;

    MailApp.sendEmail({
      to: email,
      subject: '[CG Accounts] ' + subject,
      body: 'Error occurred at: ' + new Date().toISOString() + '\n\n' +
            'Error: ' + (error.message || error) + '\n\n' +
            'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getName()
    });

  } catch (e) {
    Logger.log('Failed to send error email: ' + e.message);
  }
}

// ============ UI HELPER FUNCTIONS ============

/**
 * Gets active org info for sidebar
 */
function getActiveOrgs() {
  try {
    const config = Init.readConfig();
    return [{
      code: config.ORG_CODE,
      name: config.ORG_NAME,
      active: true
    }];
  } catch (e) {
    return [];
  }
}

/**
 * Gets recent logs for UI
 */
function getRecentLogsForUI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('RUN_LOG');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const logs = [];

    for (let i = Math.min(data.length - 1, 50); i > 0; i--) {
      logs.push({
        timestamp: data[i][0],
        sourceType: data[i][1],
        action: data[i][2],
        rowsFetched: data[i][3],
        rowsWritten: data[i][4],
        status: data[i][5],
        durationMs: data[i][6],
        errorMessage: data[i][7]
      });
    }

    return logs;
  } catch (e) {
    return [];
  }
}

/**
 * Gets settings for UI
 */
function getSettingsForUI() {
  try {
    return Init.readConfig();
  } catch (e) {
    return {};
  }
}

/**
 * Validates config from UI
 */
function validateConfigFromUI() {
  return Init.validateConfig();
}

// ============ MANUAL LEDGER CREATION ============

/**
 * Creates all ledgers - called from menu
 * This is the main manual trigger for ledger generation
 */
function createAllLedgers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Create All Ledgers',
    'This will create/update ledger sheets for all parties found in Purchase, Sales, and Bank data.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  refreshData();
}

/**
 * Prompts user to enter a specific party ID to create ledger for
 */
function createSingleLedgerPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Create Single Ledger',
    'Enter the Party ID (e.g., CG-SUP-0001):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const partyId = response.getResponseText().trim().toUpperCase();
  if (!partyId) {
    ui.alert('Error', 'Party ID cannot be empty.', ui.ButtonSet.OK);
    return;
  }

  createSingleLedger(partyId);
}

/**
 * Creates a ledger for a single party
 * @param {string} partyId - The party ID (e.g., CG-SUP-0001)
 */
function createSingleLedger(partyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();

  try {
    const config = Init.readConfig();

    // Get company info from contacts sheet using COMPANY_CONTACT_ID
    const company = getCompanyFromConfig(config);

    // Fetch party info from contacts sheet
    const partyInfo = fetchPartyFromContacts(config, partyId);
    if (!partyInfo) {
      ui.alert('Error', 'Party "' + partyId + '" not found in Contacts sheet.', ui.ButtonSet.OK);
      return;
    }

    // Fetch all transactions for this party
    const transactions = fetchTransactionsForParty(config, partyId);

    // Determine ledger category based on party ID prefix
    let ledgerCategory = 'CU';  // Default to customer
    if (partyId.includes('-SUP-')) {
      ledgerCategory = 'SU';
    }

    // Create the ledger
    const result = PartyLedger.createPartyLedger(ss, partyInfo, company, transactions, ledgerCategory);

    // Reorder tabs after creating ledger
    Init.reorderTabs(ss);

    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    ui.alert(
      'Ledger Created',
      'Party: ' + partyId + '\n' +
      'Name: ' + (partyInfo.name || 'N/A') + '\n' +
      'Type: [' + ledgerCategory + ']\n' +
      'Transactions: ' + transactions.length + '\n' +
      'Duration: ' + duration + 's',
      ui.ButtonSet.OK
    );

    Init.logRun('LEDGER', 'Create Single: ' + partyId, 0, 1, 'SUCCESS', Date.now() - startTime);

  } catch (error) {
    Init.logRun('LEDGER', 'Create Single: ' + partyId, 0, 0, 'ERROR', Date.now() - startTime, error.message);
    ui.alert('Error', 'Failed to create ledger: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Fetches party info from the Contacts sheet
 * @param {Object} config - Configuration object
 * @param {string} partyId - Party ID to find
 * @returns {Object|null} Party object or null if not found
 */
function fetchPartyFromContacts(config, partyId) {
  if (!config.CONTACTS_SHEET_ID) return null;

  try {
    const contactsSheet = SpreadsheetApp.openById(config.CONTACTS_SHEET_ID);
    const tab = contactsSheet.getSheetByName(config.CONTACTS_SHEET_NAME || 'ALL CONTACTS');
    if (!tab) return null;

    const data = tab.getDataRange().getValues();
    if (data.length < 2) return null;

    const headers = data[0].map(h => String(h).trim().toUpperCase());

    // Find column indices
    const colIdx = {
      sl: headers.indexOf('SL'),
      type: headers.indexOf('CONTACT TYPE'),
      company: headers.indexOf('COMPANY NAME'),
      addr1: headers.indexOf('ADDRESS LINE 1'),
      addr2: headers.indexOf('ADDRESS LINE 2'),
      district: headers.indexOf('DISTRICT'),
      state: headers.indexOf('STATE'),
      pin: headers.indexOf('PIN CODE'),
      gst: headers.indexOf('GST'),
      related: headers.indexOf('RELATED COMPANY'),
      contact: headers.indexOf('CONTACT PERSON'),
      mobile: headers.indexOf('MOBILE NO'),
      email: headers.indexOf('EMAIL ID')
    };

    // Find the party row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowId = String(row[colIdx.sl] || '').trim().toUpperCase();

      if (rowId === partyId) {
        // Determine party type from ID prefix
        let type = 'OTHER';
        if (partyId.includes('SUP')) type = 'SUPPLIER';
        else if (partyId.includes('CUS')) type = 'CUSTOMER';
        else if (partyId.includes('CON')) type = 'CONTRACTOR';
        else if (partyId.includes('DEA')) type = 'DEALER';
        else if (partyId.includes('REN')) type = 'RENTAL';
        else if (partyId.includes('MAS')) type = 'MASTER';

        return {
          id: partyId,
          type: type,
          name: row[colIdx.company] || '',
          address1: row[colIdx.addr1] || '',
          address2: [row[colIdx.district], row[colIdx.state], row[colIdx.pin]].filter(x => x).join(', '),
          gst: row[colIdx.gst] || '',
          phone: row[colIdx.mobile] || '',
          email: row[colIdx.email] || '',
          relatedCompany: row[colIdx.related] || '',
          contactPerson: row[colIdx.contact] || ''
        };
      }
    }

    return null;
  } catch (error) {
    Logger.log('Error fetching party from contacts: ' + error.message);
    return null;
  }
}

/**
 * Fetches all transactions for a specific party from all sources
 * @param {Object} config - Configuration object
 * @param {string} partyId - Party ID to filter by
 * @returns {Array} Array of transaction objects
 */
function fetchTransactionsForParty(config, partyId) {
  const transactions = [];

  // Fetch from Purchase
  if (config.PURCHASE_SHEET_ID) {
    const purchaseData = fetchSourceData(config, 'PURCHASE');
    if (purchaseData.data) {
      const filtered = purchaseData.data.filter(txn =>
        String(txn.partyId).toUpperCase() === partyId
      );
      transactions.push(...filtered);
    }
  }

  // Fetch from Sales
  if (config.SALES_SHEET_ID) {
    const salesData = fetchSourceData(config, 'SALES');
    if (salesData.data) {
      const filtered = salesData.data.filter(txn =>
        String(txn.partyId).toUpperCase() === partyId
      );
      transactions.push(...filtered);
    }
  }

  // Fetch from Bank (all tabs)
  if (config.BANK_SHEET_ID) {
    const bankData = fetchBankDataAllTabs(config);
    if (bankData) {
      const filtered = bankData.filter(txn =>
        String(txn.partyId).toUpperCase() === partyId
      );
      transactions.push(...filtered);
    }
  }

  // Sort by date
  transactions.sort((a, b) => {
    const dateA = a.date instanceof Date ? a.date : new Date(a.date || 0);
    const dateB = b.date instanceof Date ? b.date : new Date(b.date || 0);
    return dateA - dateB;
  });

  return transactions;
}

/**
 * Fetches bank data from all tabs that match the bank schema
 * @param {Object} config - Configuration object
 * @returns {Array} Array of transaction objects from all bank tabs
 */
function fetchBankDataAllTabs(config) {
  const transactions = [];
  const BANK_SCHEMA_HEADERS = ['DATE', 'PARTICULARS', 'DEBIT', 'CREDIT', 'BALANCE', 'L/F'];

  if (!config.BANK_SHEET_ID) return transactions;

  try {
    const bankWorkbook = SpreadsheetApp.openById(config.BANK_SHEET_ID);
    const allSheets = bankWorkbook.getSheets();

    for (const sheet of allSheets) {
      try {
        // Check row 5 for bank schema
        const row5 = sheet.getRange(5, 1, 1, 13).getValues()[0];
        const headers = row5.map(h => String(h).trim().toUpperCase());

        // Verify first 6 columns match
        const matches = BANK_SCHEMA_HEADERS.every((h, i) => headers[i] === h);
        if (!matches) continue;

        // Get data starting from row 6
        const lastRow = sheet.getLastRow();
        if (lastRow < 6) continue;

        const data = sheet.getRange(6, 1, lastRow - 5, 13).getValues();

        for (const row of data) {
          if (!row[0]) continue; // Skip empty rows

          transactions.push({
            date: row[0],
            particulars: row[1] || '',
            debit: parseNumericValue(row[2]),
            credit: parseNumericValue(row[3]),
            partyId: String(row[5] || '').trim(),
            partyName: String(row[6] || ''),
            voucherType: String(row[8] || 'BANK'),
            reference: String(row[9] || ''),
            docType: 'BANK',
            docNo: String(row[9] || '')
          });
        }
      } catch (e) {
        // Skip sheets that can't be read
      }
    }
  } catch (error) {
    Logger.log('Error fetching bank data: ' + error.message);
  }

  return transactions;
}

/**
 * Gets the company object from config for ledger headers
 * Uses COMPANY_CONTACT_ID to pull from contacts sheet if available
 * @param {Object} config - Configuration object
 * @returns {Object} Company object
 */
function getCompanyFromConfig(config) {
  // Use the Init module's fetchCompanyFromContacts for full contact lookup
  return Init.fetchCompanyFromContacts(config);
}

/**
 * Batch loads all contacts from the contacts sheet
 * Returns a map of partyId -> contact object for fast lookup
 * @param {Object} config - Configuration object
 * @returns {Object} Map of partyId to contact info
 */
function fetchAllContacts(config) {
  const contactMap = {};

  if (!config.CONTACTS_SHEET_ID) return contactMap;

  try {
    const contactsSheet = SpreadsheetApp.openById(config.CONTACTS_SHEET_ID);
    const tab = contactsSheet.getSheetByName(config.CONTACTS_SHEET_NAME || 'ALL CONTACTS');
    if (!tab) return contactMap;

    const data = tab.getDataRange().getValues();
    if (data.length < 2) return contactMap;

    const headers = data[0].map(h => String(h).trim().toUpperCase());

    // Find column indices
    const colIdx = {
      sl: headers.indexOf('SL'),
      type: headers.indexOf('CONTACT TYPE'),
      company: headers.indexOf('COMPANY NAME'),
      addr1: headers.indexOf('ADDRESS LINE 1'),
      addr2: headers.indexOf('ADDRESS LINE 2'),
      district: headers.indexOf('DISTRICT'),
      state: headers.indexOf('STATE'),
      pin: headers.indexOf('PIN CODE'),
      gst: headers.indexOf('GST'),
      related: headers.indexOf('RELATED COMPANY'),
      contact: headers.indexOf('CONTACT PERSON'),
      mobile: headers.indexOf('MOBILE NO'),
      email: headers.indexOf('EMAIL ID')
    };

    // Build contact map
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const partyId = String(row[colIdx.sl] || '').trim().toUpperCase();

      if (!partyId) continue;

      // Determine party type from ID prefix
      let type = 'OTHER';
      if (partyId.includes('SUP')) type = 'SUPPLIER';
      else if (partyId.includes('CUS')) type = 'CUSTOMER';
      else if (partyId.includes('CON')) type = 'CONTRACTOR';
      else if (partyId.includes('DEA')) type = 'DEALER';
      else if (partyId.includes('REN')) type = 'RENTAL';
      else if (partyId.includes('MAS')) type = 'MASTER';

      contactMap[partyId] = {
        id: partyId,
        type: type,
        name: row[colIdx.company] || '',
        address1: row[colIdx.addr1] || '',
        address2: [row[colIdx.district], row[colIdx.state], row[colIdx.pin]].filter(x => x).join(', '),
        gst: row[colIdx.gst] || '',
        phone: row[colIdx.mobile] || '',
        email: row[colIdx.email] || '',
        relatedCompany: row[colIdx.related] || '',
        contactPerson: row[colIdx.contact] || ''
      };
    }

  } catch (error) {
    Logger.log('Error fetching all contacts: ' + error.message);
  }

  return contactMap;
}
