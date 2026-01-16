/**
 * Configuration Module
 * Reads and manages configuration from the control sheet
 *
 * @fileoverview Configuration management for CG Accounts
 */

const Config = (function() {

  // Config sheet IDs - UPDATE THESE AFTER CREATING CONTROL SHEET
  const CONTROL_SHEET_ID = ''; // TODO: Set after creating control sheet
  const CACHE_DURATION = 300; // 5 minutes

  /**
   * Loads full configuration from control sheet
   * @returns {Object} Complete configuration object
   */
  function load() {
    // Try cache first
    const cache = CacheService.getScriptCache();
    const cached = cache.get('config');
    if (cached) {
      return JSON.parse(cached);
    }

    const config = {
      orgs: loadOrgConfig(),
      sources: loadSourceConfig(),
      columnMapping: loadColumnMapping(),
      settings: loadSettings()
    };

    // Cache for 5 minutes
    cache.put('config', JSON.stringify(config), CACHE_DURATION);

    return config;
  }

  /**
   * Loads organization configuration
   * @returns {Array} Array of org objects
   */
  function loadOrgConfig() {
    const ss = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    const sheet = ss.getSheetByName('ORG_CONFIG');
    if (!sheet) throw new Error('ORG_CONFIG sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const orgs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows

      orgs.push({
        code: row[headers.indexOf('ORG_CODE')],
        name: row[headers.indexOf('ORG_NAME')],
        ledgerSheetId: row[headers.indexOf('LEDGER_SHEET_ID')],
        active: row[headers.indexOf('ACTIVE')] === true || row[headers.indexOf('ACTIVE')] === 'TRUE'
      });
    }

    return orgs;
  }

  /**
   * Loads source configuration
   * @returns {Array} Array of source objects
   */
  function loadSourceConfig() {
    const ss = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    const sheet = ss.getSheetByName('SOURCE_CONFIG');
    if (!sheet) throw new Error('SOURCE_CONFIG sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const sources = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      sources.push({
        orgCode: row[headers.indexOf('ORG_CODE')],
        type: row[headers.indexOf('SOURCE_TYPE')],
        sheetId: row[headers.indexOf('SHEET_ID')],
        sheetName: row[headers.indexOf('SHEET_NAME')],
        active: row[headers.indexOf('ACTIVE')] === true || row[headers.indexOf('ACTIVE')] === 'TRUE'
      });
    }

    return sources;
  }

  /**
   * Loads column mapping configuration
   * @returns {Object} Mapping by source type
   */
  function loadColumnMapping() {
    const ss = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    const sheet = ss.getSheetByName('COLUMN_MAPPING');
    if (!sheet) throw new Error('COLUMN_MAPPING sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const mapping = {
      PURCHASE: {},
      SALES: {},
      BANK: {}
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const sourceType = row[headers.indexOf('SOURCE_TYPE')];
      const field = row[headers.indexOf('FIELD')];
      const columnName = row[headers.indexOf('COLUMN_NAME')];
      const required = row[headers.indexOf('REQUIRED')] === true || row[headers.indexOf('REQUIRED')] === 'TRUE';

      if (mapping[sourceType]) {
        mapping[sourceType][field] = { columnName, required };
      }
    }

    return mapping;
  }

  /**
   * Loads general settings
   * @returns {Object} Settings object
   */
  function loadSettings() {
    const ss = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    const sheet = ss.getSheetByName('SETTINGS');
    if (!sheet) {
      // Return defaults if no settings sheet
      return {
        errorEmailRecipients: [],
        logRetentionDays: 30,
        timezone: 'Asia/Kolkata'
      };
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};

    for (let i = 1; i < data.length; i++) {
      const [key, value] = data[i];
      if (key) settings[key] = value;
    }

    return settings;
  }

  /**
   * Gets the current org based on active spreadsheet
   * @returns {string|null} Org code or null
   */
  function getCurrentOrg() {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSheet) return null;

    const sheetId = activeSheet.getId();
    const config = load();

    // Check if this is a ledger sheet
    const org = config.orgs.find(o => o.ledgerSheetId === sheetId);
    if (org) return org.code;

    // Check if this is a source sheet
    const source = config.sources.find(s => s.sheetId === sheetId);
    if (source) return source.orgCode;

    return null;
  }

  /**
   * Gets source config for a specific org and type
   * @param {string} orgCode - Organization code
   * @param {string} sourceType - PURCHASE, SALES, or BANK
   * @returns {Object|null} Source config or null
   */
  function getSource(orgCode, sourceType) {
    const config = load();
    return config.sources.find(s =>
      s.orgCode === orgCode &&
      s.type === sourceType &&
      s.active
    );
  }

  /**
   * Gets org config by code
   * @param {string} orgCode - Organization code
   * @returns {Object|null} Org config or null
   */
  function getOrg(orgCode) {
    const config = load();
    return config.orgs.find(o => o.code === orgCode);
  }

  /**
   * Clears cached configuration
   */
  function clearCache() {
    CacheService.getScriptCache().remove('config');
  }

  /**
   * Validates configuration completeness
   * @returns {Object} Validation result
   */
  function validate() {
    const errors = [];
    const warnings = [];

    try {
      const config = load();

      // Check each active org has required sources
      for (const org of config.orgs.filter(o => o.active)) {
        const orgSources = config.sources.filter(s => s.orgCode === org.code && s.active);

        if (!orgSources.find(s => s.type === 'PURCHASE')) {
          warnings.push(`${org.code}: No active PURCHASE source`);
        }
        if (!orgSources.find(s => s.type === 'SALES')) {
          warnings.push(`${org.code}: No active SALES source`);
        }
        if (!orgSources.find(s => s.type === 'BANK')) {
          warnings.push(`${org.code}: No active BANK source`);
        }
        if (!org.ledgerSheetId) {
          errors.push(`${org.code}: No LEDGER_SHEET_ID configured`);
        }
      }

      // Check column mappings exist
      for (const type of ['PURCHASE', 'SALES', 'BANK']) {
        const mapping = config.columnMapping[type];
        if (!mapping || Object.keys(mapping).length === 0) {
          errors.push(`No column mapping for ${type}`);
        }
      }

    } catch (e) {
      errors.push('Failed to load config: ' + e.message);
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings
    };
  }

  // Public API
  return {
    load,
    getOrg,
    getSource,
    getCurrentOrg,
    clearCache,
    validate,
    CONTROL_SHEET_ID // Exposed for initial setup
  };

})();
