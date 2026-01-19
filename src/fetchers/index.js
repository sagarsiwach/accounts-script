/**
 * Fetchers Module (LEGACY - DEPRECATED)
 *
 * @deprecated This module is deprecated. Use fetchSourceData() and fetchBankDataAllTabs() in main.js instead.
 * This was designed for a centralized control sheet architecture.
 * Current architecture (v0.4.0+) uses per-sheet CONFIG with direct fetching in main.js
 *
 * @fileoverview Legacy data fetchers - DO NOT USE
 */

const Fetchers = (function() {

  /**
   * @deprecated Use fetchSourceData() in main.js
   * Fetches purchase data for an organization
   * @param {string} orgCode - Organization code
   * @returns {Object} Result with data array and metadata
   */
  function purchase(orgCode) {
    return fetchFromSource(orgCode, 'PURCHASE');
  }

  /**
   * @deprecated Use fetchSourceData() in main.js
   * Fetches sales data for an organization
   * @param {string} orgCode - Organization code
   * @returns {Object} Result with data array and metadata
   */
  function sales(orgCode) {
    return fetchFromSource(orgCode, 'SALES');
  }

  /**
   * Fetches bank statement data for an organization
   * @param {string} orgCode - Organization code
   * @returns {Object} Result with data array and metadata
   */
  function bank(orgCode) {
    return fetchFromSource(orgCode, 'BANK');
  }

  /**
   * Generic fetch from source sheet
   * @param {string} orgCode - Organization code
   * @param {string} sourceType - PURCHASE, SALES, or BANK
   * @returns {Object} Result object
   */
  function fetchFromSource(orgCode, sourceType) {
    const result = {
      data: [],
      rows: 0,
      status: 'pending',
      error: null
    };

    try {
      // Get source configuration
      const source = Config.getSource(orgCode, sourceType);
      if (!source) {
        result.status = 'SKIPPED';
        result.error = 'No active source configured';
        return result;
      }

      // Open source sheet
      const ss = SpreadsheetApp.openById(source.sheetId);
      const sheet = ss.getSheetByName(source.sheetName);

      if (!sheet) {
        throw new Error('Sheet not found: ' + source.sheetName);
      }

      // Get all data
      const allData = sheet.getDataRange().getValues();
      if (allData.length < 2) {
        result.status = 'SUCCESS';
        result.rows = 0;
        return result;
      }

      const headers = allData[0];
      const rows = allData.slice(1);

      // Get column mapping
      const config = Config.load();
      const mapping = config.columnMapping[sourceType];

      // Transform each row
      result.data = rows
        .filter(row => row.some(cell => cell !== '')) // Skip empty rows
        .map(row => transformRow(row, headers, mapping, sourceType));

      result.rows = result.data.length;
      result.status = 'SUCCESS';

    } catch (error) {
      result.status = 'ERROR';
      result.error = error.message;
      Logger.log('Fetch error for ' + orgCode + ' ' + sourceType + ': ' + error.message);
    }

    return result;
  }

  /**
   * Transforms a source row to standardized format
   * @param {Array} row - Source data row
   * @param {Array} headers - Column headers
   * @param {Object} mapping - Column mapping config
   * @param {string} sourceType - Source type
   * @returns {Object} Transformed entry
   */
  function transformRow(row, headers, mapping, sourceType) {
    const entry = {
      date: null,
      docType: sourceType,
      docNo: '',
      accountCode: '',
      partyId: '',
      partyName: '',
      gstNumber: '',
      debit: 0,
      credit: 0,
      igst: 0,
      cgst: 0,
      sgst: 0,
      gstTotal: 0,
      narration: '',
      sourceRef: '',
      raw: {} // Keep raw values for debugging
    };

    // Apply each mapping
    for (const [field, config] of Object.entries(mapping)) {
      const colIndex = Utils.getColumnIndex(headers, config.columnName);
      if (colIndex === -1) continue;

      const value = row[colIndex];
      entry.raw[field] = value;

      switch (field) {
        case 'DATE':
          entry.date = Utils.parseDate(value);
          break;

        case 'DOC_NO':
        case 'INVOICE_NO':
          entry.docNo = String(value || '');
          break;

        case 'PARTY_ID':
        case 'L/F':
          entry.partyId = String(value || '');
          break;

        case 'PARTY_NAME':
        case 'SUPPLIER_NAME':
        case 'CUSTOMER_NAME':
          entry.partyName = String(value || '');
          break;

        case 'GST_NUMBER':
          entry.gstNumber = String(value || '');
          break;

        case 'AMOUNT':
        case 'GRAND_TOTAL':
          const amount = Utils.parseNumber(value);
          if (sourceType === 'PURCHASE') {
            entry.credit = amount; // Purchases create liability
          } else if (sourceType === 'SALES') {
            entry.debit = amount; // Sales create receivable
          }
          break;

        case 'DEBIT':
          entry.debit = Utils.parseNumber(value);
          break;

        case 'CREDIT':
          entry.credit = Utils.parseNumber(value);
          break;

        case 'IGST':
          entry.igst = Utils.parseNumber(value);
          break;

        case 'CGST':
          entry.cgst = Utils.parseNumber(value);
          break;

        case 'SGST':
          entry.sgst = Utils.parseNumber(value);
          break;

        case 'GST_TOTAL':
          entry.gstTotal = Utils.parseNumber(value);
          break;

        case 'ACCOUNT':
        case 'ACCOUNT_CODE':
          entry.accountCode = String(value || '');
          break;

        case 'NARRATION':
        case 'REMARKS':
        case 'PARTICULARS':
          entry.narration = String(value || '');
          break;
      }
    }

    // Calculate GST total if not provided
    if (!entry.gstTotal && (entry.igst || entry.cgst || entry.sgst)) {
      entry.gstTotal = entry.igst + entry.cgst + entry.sgst;
    }

    return entry;
  }

  // Public API
  return {
    purchase,
    sales,
    bank
  };

})();
