/**
 * Unit tests for Ledgers module
 */

// Mock Google Apps Script globals
global.Logger = { log: jest.fn() };
global.SpreadsheetApp = {
  openById: jest.fn()
};

// Mock Config module
global.Config = {
  getOrg: jest.fn()
};

// Mock Utils module
global.Utils = {
  chunk: (arr, size) => {
    const chunks = [];
    for (let i = 0; i < arr.length; i += size) {
      chunks.push(arr.slice(i, i + size));
    }
    return chunks;
  }
};

// Load the module
const ledgersCode = require('fs').readFileSync('./src/ledgers/index.js', 'utf8');
eval(ledgersCode);

describe('Ledgers', () => {

  describe('generateMasterEntries', () => {
    test('combines entries from all sources', () => {
      const sourceData = {
        purchase: [
          { date: new Date('2026-01-15'), partyId: 'SUP-001', credit: 1000 }
        ],
        sales: [
          { date: new Date('2026-01-16'), partyId: 'CUST-001', debit: 500 }
        ],
        bank: [
          { date: new Date('2026-01-14'), partyId: 'PARTY-001', debit: 200, credit: 0 }
        ]
      };

      const entries = Ledgers.generateMasterEntries(sourceData);

      expect(entries.length).toBe(3);
      // Should be sorted by date
      expect(entries[0].date.getDate()).toBe(14);
      expect(entries[1].date.getDate()).toBe(15);
      expect(entries[2].date.getDate()).toBe(16);
    });

    test('handles empty source data', () => {
      const sourceData = {
        purchase: [],
        sales: [],
        bank: []
      };

      const entries = Ledgers.generateMasterEntries(sourceData);
      expect(entries).toEqual([]);
    });

    test('handles null source data', () => {
      const sourceData = {
        purchase: null,
        sales: undefined,
        bank: []
      };

      const entries = Ledgers.generateMasterEntries(sourceData);
      expect(entries).toEqual([]);
    });
  });

  describe('groupByParty', () => {
    test('groups entries by party and type', () => {
      const entries = [
        { partyId: 'SUP-001', partyName: 'Supplier 1', docType: 'PURCHASE', credit: 1000 },
        { partyId: 'SUP-001', partyName: 'Supplier 1', docType: 'PURCHASE', credit: 500 },
        { partyId: 'CUST-001', partyName: 'Customer 1', docType: 'SALES', debit: 2000 }
      ];

      const grouped = Ledgers.groupByParty(entries);

      expect(Object.keys(grouped.suppliers)).toContain('SUP-001');
      expect(grouped.suppliers['SUP-001'].entries.length).toBe(2);
      expect(Object.keys(grouped.customers)).toContain('CUST-001');
      expect(grouped.customers['CUST-001'].entries.length).toBe(1);
    });

    test('skips entries without partyId', () => {
      const entries = [
        { partyId: '', partyName: '', docType: 'PURCHASE', credit: 1000 },
        { partyId: 'SUP-001', partyName: 'Supplier', docType: 'PURCHASE', credit: 500 }
      ];

      const grouped = Ledgers.groupByParty(entries);

      expect(Object.keys(grouped.suppliers).length).toBe(1);
    });
  });

  describe('generate', () => {
    test('returns error when no ledger sheet configured', () => {
      Config.getOrg.mockReturnValue({ code: 'CM', ledgerSheetId: '' });

      const result = Ledgers.generate('CM', { purchase: [], sales: [], bank: [] });

      expect(result.status).toBe('ERROR');
      expect(result.error).toContain('No ledger sheet configured');
    });

    test('writes to ledger master sheet', () => {
      const mockSheet = {
        getLastRow: () => 0,
        getLastColumn: () => 10,
        appendRow: jest.fn(),
        setFrozenRows: jest.fn(),
        getRange: () => ({
          setValues: jest.fn(),
          setFontWeight: jest.fn(),
          setBackground: jest.fn(),
          setNumberFormat: jest.fn(),
          getValues: () => [['DATE', 'DEBIT', 'CREDIT', 'BALANCE']],
          clear: jest.fn()
        }),
        autoResizeColumn: jest.fn()
      };

      const mockSpreadsheet = {
        getSheetByName: (name) => name === 'Ledger Master' ? mockSheet : null,
        insertSheet: () => mockSheet
      };

      SpreadsheetApp.openById.mockReturnValue(mockSpreadsheet);
      Config.getOrg.mockReturnValue({ code: 'CM', ledgerSheetId: 'test-id' });

      const sourceData = {
        purchase: [
          { date: new Date('2026-01-16'), partyId: 'SUP-001', partyName: 'Test', credit: 1000 }
        ],
        sales: [],
        bank: []
      };

      const result = Ledgers.generate('CM', sourceData);

      expect(result.status).toBe('SUCCESS');
      expect(result.rows).toBeGreaterThan(0);
    });
  });

});
