/**
 * Classic Group Accounts Script
 * Main entry point - menu and trigger handlers
 *
 * @fileoverview Entry points for the accounts aggregation system
 * @author Sagar Siwach
 * @version 0.1.0
 */

/**
 * Creates the custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CG Accounts')
    .addItem('Open Dashboard', 'showSidebar')
    .addSeparator()
    .addItem('Refresh All Orgs', 'refreshAllOrgs')
    .addItem('Refresh Current Org', 'refreshCurrentOrg')
    .addSeparator()
    .addSubMenu(ui.createMenu('Quick Actions')
      .addItem('View Logs', 'showLogs')
      .addItem('Test Connection', 'testConnection'))
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
}

/**
 * Installable trigger for hourly refresh
 */
function hourlyRefresh() {
  const startTime = new Date();
  Logger.log('Starting hourly refresh at ' + startTime.toISOString());

  try {
    const result = refreshAllOrgs();
    Logger.log('Hourly refresh completed: ' + JSON.stringify(result));
  } catch (error) {
    Logger.log('Hourly refresh failed: ' + error.message);
    sendErrorEmail('Hourly Refresh Failed', error);
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
 * Shows settings dialog
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('ui/Settings')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

/**
 * Refreshes all active organizations
 * @returns {Object} Summary of refresh operation
 */
function refreshAllOrgs() {
  const config = getConfig();
  const results = {
    success: [],
    failed: [],
    skipped: [],
    startTime: new Date(),
    endTime: null
  };

  const activeOrgs = config.orgs.filter(org => org.active);

  for (const org of activeOrgs) {
    try {
      const orgResult = refreshOrg(org.code);
      results.success.push({ org: org.code, ...orgResult });
    } catch (error) {
      results.failed.push({ org: org.code, error: error.message });
      logError(org.code, 'REFRESH', error);
    }
  }

  results.endTime = new Date();
  results.duration = results.endTime - results.startTime;

  // Send email if any failures
  if (results.failed.length > 0) {
    sendErrorEmail('Refresh Errors', results.failed);
  }

  return results;
}

/**
 * Refreshes a single organization
 * @param {string} orgCode - Organization code (CM, KM, DEV, CPI, SM)
 * @returns {Object} Result of refresh operation
 */
function refreshOrg(orgCode) {
  const startTime = new Date();
  const result = {
    purchase: { rows: 0, status: 'pending' },
    sales: { rows: 0, status: 'pending' },
    bank: { rows: 0, status: 'pending' },
    ledger: { rows: 0, status: 'pending' }
  };

  // Fetch from sources
  result.purchase = fetchPurchaseData(orgCode);
  result.sales = fetchSalesData(orgCode);
  result.bank = fetchBankData(orgCode);

  // Generate ledgers
  result.ledger = generateLedgers(orgCode, {
    purchase: result.purchase.data,
    sales: result.sales.data,
    bank: result.bank.data
  });

  // Log the run
  logRun(orgCode, result, startTime);

  return result;
}

/**
 * Refreshes the current org based on active spreadsheet
 */
function refreshCurrentOrg() {
  const currentOrgCode = getCurrentOrgCode();
  if (!currentOrgCode) {
    SpreadsheetApp.getUi().alert('Could not determine current organization. Please use Refresh All or open an org-specific sheet.');
    return;
  }

  const result = refreshOrg(currentOrgCode);
  SpreadsheetApp.getUi().alert('Refresh complete for ' + currentOrgCode + '\n\nPurchase: ' + result.purchase.rows + ' rows\nSales: ' + result.sales.rows + ' rows\nBank: ' + result.bank.rows + ' rows');
}

/**
 * Tests connection to all configured sheets
 */
function testConnection() {
  const config = getConfig();
  const results = [];

  for (const org of config.orgs) {
    if (!org.active) continue;

    const sources = config.sources.filter(s => s.orgCode === org.code && s.active);
    for (const source of sources) {
      try {
        const ss = SpreadsheetApp.openById(source.sheetId);
        const sheet = ss.getSheetByName(source.sheetName);
        if (sheet) {
          results.push({ org: org.code, source: source.type, status: 'OK' });
        } else {
          results.push({ org: org.code, source: source.type, status: 'Sheet not found: ' + source.sheetName });
        }
      } catch (e) {
        results.push({ org: org.code, source: source.type, status: 'Error: ' + e.message });
      }
    }
  }

  // Show results
  let message = 'Connection Test Results:\n\n';
  for (const r of results) {
    message += r.org + ' - ' + r.source + ': ' + r.status + '\n';
  }
  SpreadsheetApp.getUi().alert(message);
}

// Placeholder functions - implemented in respective modules
function getConfig() { return Config.load(); }
function getCurrentOrgCode() { return Config.getCurrentOrg(); }
function fetchPurchaseData(orgCode) { return Fetchers.purchase(orgCode); }
function fetchSalesData(orgCode) { return Fetchers.sales(orgCode); }
function fetchBankData(orgCode) { return Fetchers.bank(orgCode); }
function generateLedgers(orgCode, data) { return Ledgers.generate(orgCode, data); }
function logRun(orgCode, result, startTime) { return AppLogger.logRun(orgCode, result, startTime); }
function logError(orgCode, operation, error) { return AppLogger.logError(orgCode, operation, error); }
function sendErrorEmail(subject, errors) { return AppLogger.sendErrorEmail(subject, errors); }
