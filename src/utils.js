/**
 * Utility Functions
 * Common helper functions used across modules
 *
 * @fileoverview Shared utilities for CG Accounts
 */

const Utils = (function() {

  /**
   * Gets column index by header name
   * @param {Array} headers - Header row
   * @param {string} columnName - Column name to find
   * @returns {number} Column index or -1
   */
  function getColumnIndex(headers, columnName) {
    return headers.findIndex(h =>
      String(h).trim().toUpperCase() === String(columnName).trim().toUpperCase()
    );
  }

  /**
   * Extracts a value from a row by column name
   * @param {Array} row - Data row
   * @param {Array} headers - Header row
   * @param {string} columnName - Column name
   * @param {*} defaultValue - Default if not found
   * @returns {*} Value or default
   */
  function getValue(row, headers, columnName, defaultValue = null) {
    const index = getColumnIndex(headers, columnName);
    if (index === -1) return defaultValue;
    return row[index] !== undefined && row[index] !== '' ? row[index] : defaultValue;
  }

  /**
   * Parses a date value safely
   * @param {*} value - Value to parse
   * @returns {Date|null} Parsed date or null
   */
  function parseDate(value) {
    if (!value) return null;
    if (value instanceof Date) return value;

    // Try parsing string
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  /**
   * Parses a number value safely
   * @param {*} value - Value to parse
   * @returns {number} Parsed number or 0
   */
  function parseNumber(value) {
    if (!value) return 0;
    if (typeof value === 'number') return value;

    // Remove currency symbols and commas
    const cleaned = String(value).replace(/[₹$,\s]/g, '');
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : parsed;
  }

  /**
   * Formats a date for display
   * @param {Date} date - Date to format
   * @param {string} format - Format string (default: DD-MM-YYYY)
   * @returns {string} Formatted date
   */
  function formatDate(date, format = 'DD-MM-YYYY') {
    if (!date || !(date instanceof Date)) return '';

    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yyyy = date.getFullYear();
    const hh = String(date.getHours()).padStart(2, '0');
    const min = String(date.getMinutes()).padStart(2, '0');

    return format
      .replace('DD', dd)
      .replace('MM', mm)
      .replace('YYYY', yyyy)
      .replace('HH', hh)
      .replace('mm', min);
  }

  /**
   * Formats a number as currency
   * @param {number} value - Number to format
   * @param {string} currency - Currency symbol (default: ₹)
   * @returns {string} Formatted currency
   */
  function formatCurrency(value, currency = '₹') {
    if (value === null || value === undefined) return '';
    const num = parseNumber(value);
    return currency + ' ' + num.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  /**
   * Generates a unique ID
   * @param {string} prefix - ID prefix
   * @returns {string} Unique ID
   */
  function generateId(prefix = '') {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 8);
    return prefix + timestamp + random;
  }

  /**
   * Chunks an array into smaller arrays
   * @param {Array} array - Array to chunk
   * @param {number} size - Chunk size
   * @returns {Array} Array of chunks
   */
  function chunk(array, size) {
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  }

  /**
   * Deep clones an object
   * @param {Object} obj - Object to clone
   * @returns {Object} Cloned object
   */
  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  /**
   * Validates required fields in an object
   * @param {Object} obj - Object to validate
   * @param {Array} requiredFields - List of required field names
   * @returns {Object} Validation result
   */
  function validateRequired(obj, requiredFields) {
    const missing = [];
    for (const field of requiredFields) {
      if (obj[field] === undefined || obj[field] === null || obj[field] === '') {
        missing.push(field);
      }
    }
    return {
      valid: missing.length === 0,
      missing
    };
  }

  /**
   * Transforms source data row to standardized ledger format
   * @param {Array} row - Source data row
   * @param {Array} headers - Source headers
   * @param {Object} mapping - Column mapping
   * @param {string} docType - Document type (PURCHASE, SALES, BANK)
   * @returns {Object} Standardized ledger entry
   */
  function transformToLedgerEntry(row, headers, mapping, docType) {
    const entry = {
      date: null,
      docType: docType,
      docNo: '',
      accountCode: '',
      partyId: '',
      partyName: '',
      debit: 0,
      credit: 0,
      narration: '',
      sourceRef: ''
    };

    // Apply mapping
    for (const [field, config] of Object.entries(mapping)) {
      const value = getValue(row, headers, config.columnName);

      switch (field) {
        case 'DATE':
          entry.date = parseDate(value);
          break;
        case 'DOC_NO':
          entry.docNo = String(value || '');
          break;
        case 'PARTY_ID':
          entry.partyId = String(value || '');
          break;
        case 'PARTY_NAME':
          entry.partyName = String(value || '');
          break;
        case 'AMOUNT':
        case 'GRAND_TOTAL':
          // For purchases: credit to supplier
          // For sales: debit from customer
          const amount = parseNumber(value);
          if (docType === 'PURCHASE') {
            entry.credit = amount;
          } else if (docType === 'SALES') {
            entry.debit = amount;
          }
          break;
        case 'DEBIT':
          entry.debit = parseNumber(value);
          break;
        case 'CREDIT':
          entry.credit = parseNumber(value);
          break;
        case 'NARRATION':
        case 'REMARKS':
          entry.narration = String(value || '');
          break;
      }
    }

    return entry;
  }

  // Public API
  return {
    getColumnIndex,
    getValue,
    parseDate,
    parseNumber,
    formatDate,
    formatCurrency,
    generateId,
    chunk,
    deepClone,
    validateRequired,
    transformToLedgerEntry
  };

})();
