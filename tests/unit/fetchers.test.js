/**
 * Unit tests for Fetchers module
 */

// Mock Google Apps Script globals
global.Logger = { log: jest.fn() };
global.SpreadsheetApp = {
  openById: jest.fn()
};
global.CacheService = {
  getScriptCache: () => ({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn()
  })
};

// Mock Config module
global.Config = {
  load: jest.fn(),
  getSource: jest.fn(),
  getOrg: jest.fn()
};

// Mock Utils module
global.Utils = {
  getColumnIndex: (headers, name) => headers.findIndex(h => h === name),
  parseDate: (val) => val ? new Date(val) : null,
  parseNumber: (val) => parseFloat(val) || 0
};

// Load the module (replace const with global assignment for Jest compatibility)
const fetchersCode = require('fs').readFileSync('./src/fetchers/index.js', 'utf8');
eval(fetchersCode.replace('const Fetchers =', 'global.Fetchers ='));
const Fetchers = global.Fetchers;

describe('Fetchers', () => {

  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('purchase', () => {
    test('returns skipped when no source configured', () => {
      Config.getSource.mockReturnValue(null);

      const result = Fetchers.purchase('CM');

      expect(result.status).toBe('SKIPPED');
      expect(result.error).toContain('No active source');
    });

    test('fetches and transforms data correctly', () => {
      // Setup mock source config
      Config.getSource.mockReturnValue({
        sheetId: 'test-sheet-id',
        sheetName: 'Purchase Register'
      });

      Config.load.mockReturnValue({
        columnMapping: {
          PURCHASE: {
            DATE: { columnName: 'INW_DATE' },
            PARTY_ID: { columnName: 'L/F' },
            PARTY_NAME: { columnName: 'SUPPLIER_NAME' },
            AMOUNT: { columnName: 'GRAND_TOTAL' }
          }
        }
      });

      // Mock spreadsheet
      const mockSheet = {
        getDataRange: () => ({
          getValues: () => [
            ['INW_DATE', 'L/F', 'SUPPLIER_NAME', 'GRAND_TOTAL'],
            ['2026-01-16', 'SUP-001', 'Test Supplier', 10000]
          ]
        })
      };

      SpreadsheetApp.openById.mockReturnValue({
        getSheetByName: () => mockSheet
      });

      const result = Fetchers.purchase('CM');

      expect(result.status).toBe('SUCCESS');
      expect(result.rows).toBe(1);
      expect(result.data[0].partyId).toBe('SUP-001');
      expect(result.data[0].credit).toBe(10000); // Purchases are credits
    });

    test('handles sheet not found error', () => {
      Config.getSource.mockReturnValue({
        sheetId: 'test-sheet-id',
        sheetName: 'Missing Sheet'
      });

      SpreadsheetApp.openById.mockReturnValue({
        getSheetByName: () => null
      });

      const result = Fetchers.purchase('CM');

      expect(result.status).toBe('ERROR');
      expect(result.error).toContain('Sheet not found');
    });
  });

  describe('sales', () => {
    test('marks amounts as debits for sales', () => {
      Config.getSource.mockReturnValue({
        sheetId: 'test-sheet-id',
        sheetName: 'Sales Register'
      });

      Config.load.mockReturnValue({
        columnMapping: {
          SALES: {
            DATE: { columnName: 'INVOICE_DATE' },
            PARTY_ID: { columnName: 'L/F' },
            AMOUNT: { columnName: 'GRAND_TOTAL' }
          }
        }
      });

      const mockSheet = {
        getDataRange: () => ({
          getValues: () => [
            ['INVOICE_DATE', 'L/F', 'GRAND_TOTAL'],
            ['2026-01-16', 'CUST-001', 5000]
          ]
        })
      };

      SpreadsheetApp.openById.mockReturnValue({
        getSheetByName: () => mockSheet
      });

      const result = Fetchers.sales('CM');

      expect(result.status).toBe('SUCCESS');
      expect(result.data[0].debit).toBe(5000); // Sales are debits
      expect(result.data[0].credit).toBe(0);
    });
  });

  describe('bank', () => {
    test('handles debit and credit columns separately', () => {
      Config.getSource.mockReturnValue({
        sheetId: 'test-sheet-id',
        sheetName: 'Bank Statement'
      });

      Config.load.mockReturnValue({
        columnMapping: {
          BANK: {
            DATE: { columnName: 'DATE' },
            PARTY_ID: { columnName: 'L/F' },
            DEBIT: { columnName: 'DEBIT' },
            CREDIT: { columnName: 'CREDIT' }
          }
        }
      });

      const mockSheet = {
        getDataRange: () => ({
          getValues: () => [
            ['DATE', 'L/F', 'DEBIT', 'CREDIT'],
            ['2026-01-16', 'PARTY-001', 1000, ''],
            ['2026-01-17', 'PARTY-002', '', 2000]
          ]
        })
      };

      SpreadsheetApp.openById.mockReturnValue({
        getSheetByName: () => mockSheet
      });

      const result = Fetchers.bank('CM');

      expect(result.status).toBe('SUCCESS');
      expect(result.rows).toBe(2);
      expect(result.data[0].debit).toBe(1000);
      expect(result.data[1].credit).toBe(2000);
    });
  });

});
