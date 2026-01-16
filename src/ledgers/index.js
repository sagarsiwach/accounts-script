/**
 * Ledgers Module
 * Generates Ledger Master and individual party ledgers
 *
 * @fileoverview Ledger generation for CG Accounts
 */

const Ledgers = (function() {

  /**
   * Generates all ledgers for an organization
   * @param {string} orgCode - Organization code
   * @param {Object} sourceData - Data from fetchers {purchase, sales, bank}
   * @returns {Object} Result with rows written and status
   */
  function generate(orgCode, sourceData) {
    const result = {
      rows: 0,
      status: 'pending',
      error: null,
      details: {
        master: 0,
        supplier: 0,
        customer: 0,
        contractor: 0
      }
    };

    try {
      // Get ledger sheet for this org
      const org = Config.getOrg(orgCode);
      if (!org || !org.ledgerSheetId) {
        throw new Error('No ledger sheet configured for ' + orgCode);
      }

      const ss = SpreadsheetApp.openById(org.ledgerSheetId);

      // Step 1: Generate Ledger Master
      const masterEntries = generateMasterEntries(sourceData);
      result.details.master = writeLedgerMaster(ss, masterEntries);

      // Step 2: Generate individual party ledgers
      const partyLedgers = groupByParty(masterEntries);

      // Supplier Ledgers (from purchases)
      result.details.supplier = writePartyLedgers(ss, partyLedgers.suppliers, 'Supplier');

      // Customer Ledgers (from sales)
      result.details.customer = writePartyLedgers(ss, partyLedgers.customers, 'Customer');

      // Contractor Ledgers (if any - usually from bank/expenses)
      result.details.contractor = writePartyLedgers(ss, partyLedgers.contractors, 'Contractor');

      result.rows = result.details.master;
      result.status = 'SUCCESS';

    } catch (error) {
      result.status = 'ERROR';
      result.error = error.message;
      Logger.log('Ledger generation error for ' + orgCode + ': ' + error.message);
    }

    return result;
  }

  /**
   * Combines all source data into master ledger entries
   * @param {Object} sourceData - Data from all sources
   * @returns {Array} Combined entries sorted by date
   */
  function generateMasterEntries(sourceData) {
    const entries = [];

    // Add purchase entries
    if (sourceData.purchase && sourceData.purchase.length > 0) {
      sourceData.purchase.forEach(entry => {
        entries.push({
          ...entry,
          docType: 'PURCHASE',
          ledgerType: 'supplier'
        });
      });
    }

    // Add sales entries
    if (sourceData.sales && sourceData.sales.length > 0) {
      sourceData.sales.forEach(entry => {
        entries.push({
          ...entry,
          docType: 'SALES',
          ledgerType: 'customer'
        });
      });
    }

    // Add bank entries
    if (sourceData.bank && sourceData.bank.length > 0) {
      sourceData.bank.forEach(entry => {
        entries.push({
          ...entry,
          docType: 'BANK',
          ledgerType: entry.credit > 0 ? 'supplier' : 'customer' // Rough classification
        });
      });
    }

    // Sort by date
    entries.sort((a, b) => {
      const dateA = a.date instanceof Date ? a.date : new Date(0);
      const dateB = b.date instanceof Date ? b.date : new Date(0);
      return dateA - dateB;
    });

    return entries;
  }

  /**
   * Writes entries to the Ledger Master sheet
   * @param {Spreadsheet} ss - Target spreadsheet
   * @param {Array} entries - Ledger entries
   * @returns {number} Rows written
   */
  function writeLedgerMaster(ss, entries) {
    let sheet = ss.getSheetByName('Ledger Master');

    // Create sheet if doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Ledger Master');
    }

    // Clear existing data (except header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
    }

    // Write headers if needed
    if (sheet.getLastRow() === 0) {
      const headers = [
        'DATE', 'DOC_TYPE', 'DOC_NO', 'ACCOUNT_CODE', 'PARTY_ID', 'PARTY_NAME',
        'DEBIT', 'CREDIT', 'BALANCE', 'IGST', 'CGST', 'SGST', 'GST_TOTAL', 'NARRATION'
      ];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);

      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
    }

    if (entries.length === 0) return 0;

    // Prepare data rows
    let runningBalance = 0;
    const dataRows = entries.map(entry => {
      runningBalance += (entry.debit || 0) - (entry.credit || 0);

      return [
        entry.date,
        entry.docType,
        entry.docNo,
        entry.accountCode,
        entry.partyId,
        entry.partyName,
        entry.debit || '',
        entry.credit || '',
        runningBalance,
        entry.igst || '',
        entry.cgst || '',
        entry.sgst || '',
        entry.gstTotal || '',
        entry.narration
      ];
    });

    // Write in batches for performance
    const BATCH_SIZE = 500;
    const batches = Utils.chunk(dataRows, BATCH_SIZE);
    let currentRow = 2;

    for (const batch of batches) {
      const range = sheet.getRange(currentRow, 1, batch.length, batch[0].length);
      range.setValues(batch);
      currentRow += batch.length;
    }

    // Format columns
    formatLedgerSheet(sheet);

    return entries.length;
  }

  /**
   * Groups entries by party for individual ledgers
   * @param {Array} entries - All ledger entries
   * @returns {Object} Grouped entries {suppliers, customers, contractors}
   */
  function groupByParty(entries) {
    const grouped = {
      suppliers: {},
      customers: {},
      contractors: {}
    };

    for (const entry of entries) {
      if (!entry.partyId) continue;

      const partyKey = entry.partyId;
      const partyData = {
        id: entry.partyId,
        name: entry.partyName,
        entries: []
      };

      if (entry.docType === 'PURCHASE') {
        if (!grouped.suppliers[partyKey]) {
          grouped.suppliers[partyKey] = { ...partyData };
        }
        grouped.suppliers[partyKey].entries.push(entry);
      } else if (entry.docType === 'SALES') {
        if (!grouped.customers[partyKey]) {
          grouped.customers[partyKey] = { ...partyData };
        }
        grouped.customers[partyKey].entries.push(entry);
      }
      // Bank entries could go to either based on type
    }

    return grouped;
  }

  /**
   * Writes individual party ledger sheets
   * @param {Spreadsheet} ss - Target spreadsheet
   * @param {Object} partyData - Grouped party data
   * @param {string} type - 'Supplier', 'Customer', or 'Contractor'
   * @returns {number} Total rows written
   */
  function writePartyLedgers(ss, partyData, type) {
    let totalRows = 0;
    const parties = Object.values(partyData);

    for (const party of parties) {
      if (!party.entries || party.entries.length === 0) continue;

      // Create sheet name (max 100 chars, sanitized)
      const sheetName = sanitizeSheetName(type + ' - ' + (party.name || party.id));

      let sheet = ss.getSheetByName(sheetName);

      // Create or clear sheet
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      // Write header
      const headers = ['DATE', 'DOC_TYPE', 'DOC_NO', 'DEBIT', 'CREDIT', 'BALANCE', 'NARRATION'];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);

      // Format header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');

      // Prepare data with running balance
      let runningBalance = 0; // TODO: Add opening balance support
      const dataRows = party.entries.map(entry => {
        runningBalance += (entry.debit || 0) - (entry.credit || 0);
        return [
          entry.date,
          entry.docType,
          entry.docNo,
          entry.debit || '',
          entry.credit || '',
          runningBalance,
          entry.narration
        ];
      });

      // Write data
      if (dataRows.length > 0) {
        const range = sheet.getRange(2, 1, dataRows.length, dataRows[0].length);
        range.setValues(dataRows);
        totalRows += dataRows.length;
      }

      // Format
      formatLedgerSheet(sheet);
    }

    return totalRows;
  }

  /**
   * Formats a ledger sheet with standard styling
   * @param {Sheet} sheet - Sheet to format
   */
  function formatLedgerSheet(sheet) {
    // Auto-resize columns
    const lastCol = sheet.getLastColumn();
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
    }

    // Format date column
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('dd-mm-yyyy');
    }

    // Format number columns (Debit, Credit, Balance)
    // Find these columns by header
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const numberCols = ['DEBIT', 'CREDIT', 'BALANCE', 'IGST', 'CGST', 'SGST', 'GST_TOTAL'];

    numberCols.forEach(colName => {
      const colIndex = headers.indexOf(colName);
      if (colIndex !== -1 && lastRow > 1) {
        sheet.getRange(2, colIndex + 1, lastRow - 1, 1).setNumberFormat('#,##0.00');
      }
    });
  }

  /**
   * Sanitizes a string for use as a sheet name
   * @param {string} name - Raw name
   * @returns {string} Sanitized name
   */
  function sanitizeSheetName(name) {
    // Remove invalid characters
    let sanitized = String(name)
      .replace(/[\/\\?*\[\]:]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();

    // Truncate to 100 chars
    if (sanitized.length > 100) {
      sanitized = sanitized.substring(0, 97) + '...';
    }

    return sanitized || 'Unnamed';
  }

  // Public API
  return {
    generate,
    generateMasterEntries,
    groupByParty
  };

})();
