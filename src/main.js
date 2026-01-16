/**
 * Classic Group Accounts Script
 * Main entry point - menu and trigger handlers
 *
 * @fileoverview Entry points for the accounts aggregation system
 * @author Sagar Siwach
 * @version 0.2.0
 *
 * Architecture: Each org's Ledger sheet is self-contained with:
 * - CONFIG tab: All settings (sources, mappings)
 * - Ledger Master tab: Index of all party ledgers with hyperlinks
 * - RUN_LOG tab: Execution history
 * - Individual party ledger tabs (CG-SUP-0001, CG-CUST-0001, etc.)
 */

/**
 * Creates the custom menu when spreadsheet opens
 * Works for both container-bound and add-on installations
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('CG Accounts')
    .addItem('Open Dashboard', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Initialize Sheet', 'initializeSheet')
      .addItem('Test Connection', 'testConnection'))
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

    // Fetch from sources
    ui.showSidebar(HtmlService.createHtmlOutput('<p>Fetching Purchase data...</p>').setTitle('Refreshing'));

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
      result.bank = fetchSourceData(config, 'BANK');
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

    const headers = allData[0];
    const rows = allData.slice(1);

    // Get column mappings for this source type
    const mapping = getColumnMapping(config, sourceType);

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
 */
function getColumnMapping(config, sourceType) {
  const prefix = sourceType === 'PURCHASE' ? 'PUR_' :
                 sourceType === 'SALES' ? 'SAL_' : 'BANK_';

  return {
    date: config[prefix + 'DATE_COL'],
    partyId: config[prefix + 'PARTY_ID_COL'],
    partyName: config[prefix + 'PARTY_NAME_COL'],
    invoice: config[prefix + 'INVOICE_COL'],
    amount: config[prefix + 'AMOUNT_COL'],
    gst: config[prefix + 'GST_COL'],
    debit: config[prefix + 'DEBIT_COL'] || config['BANK_DEBIT_COL'],
    credit: config[prefix + 'CREDIT_COL'] || config['BANK_CREDIT_COL'],
    particulars: config[prefix + 'PARTICULARS_COL'] || config['BANK_PARTICULARS_COL'],
    reference: config[prefix + 'REF_COL'] || config['BANK_REF_COL'],
    voucherType: config[prefix + 'VOUCHER_COL'] || config['BANK_VOUCHER_COL'],
    remarks: config[prefix + 'REMARKS_COL']
  };
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
 */
function generateAllLedgers(ss, config, transactions) {
  const result = { count: 0, status: 'pending', error: null };

  try {
    // Group transactions by party
    const partyMap = {};

    for (const txn of transactions) {
      if (!txn.partyId) continue;

      if (!partyMap[txn.partyId]) {
        partyMap[txn.partyId] = {
          id: txn.partyId,
          name: txn.partyName,
          type: txn.docType === 'PURCHASE' ? 'SUPPLIER' :
                txn.docType === 'SALES' ? 'CUSTOMER' : 'OTHER',
          transactions: [],
          totalDebit: 0,
          totalCredit: 0,
          lastTransaction: null
        };
      }

      partyMap[txn.partyId].transactions.push(txn);
      partyMap[txn.partyId].totalDebit += txn.debit || 0;
      partyMap[txn.partyId].totalCredit += txn.credit || 0;

      if (!partyMap[txn.partyId].lastTransaction ||
          txn.date > partyMap[txn.partyId].lastTransaction) {
        partyMap[txn.partyId].lastTransaction = txn.date;
      }
    }

    const parties = Object.values(partyMap);

    // Create individual ledger sheets
    for (const party of parties) {
      // Sort transactions by date
      party.transactions.sort((a, b) => {
        const dateA = a.date instanceof Date ? a.date : new Date(a.date || 0);
        const dateB = b.date instanceof Date ? b.date : new Date(b.date || 0);
        return dateA - dateB;
      });

      PartyLedger.createPartyLedger(ss, party, party.transactions);
      result.count++;
    }

    // Update Ledger Master index
    PartyLedger.updateLedgerMasterIndex(ss, parties);

    result.status = 'SUCCESS';

  } catch (error) {
    result.status = 'ERROR';
    result.error = error.message;
    Logger.log('Ledger generation error: ' + error.message);
  }

  return result;
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
