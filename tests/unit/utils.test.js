/**
 * Unit tests for Utils module
 */

// Mock Google Apps Script globals
global.Logger = { log: jest.fn() };
global.CacheService = {
  getScriptCache: () => ({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn()
  })
};

// Import the module (we need to handle the IIFE pattern)
const utilsCode = require('fs').readFileSync('./src/utils.js', 'utf8');
// Execute and assign to global since GAS uses global scope
eval(utilsCode.replace('const Utils =', 'global.Utils ='));
const Utils = global.Utils;

describe('Utils', () => {

  describe('getColumnIndex', () => {
    const headers = ['DATE', 'NAME', 'AMOUNT', 'Status'];

    test('finds exact match', () => {
      expect(Utils.getColumnIndex(headers, 'NAME')).toBe(1);
    });

    test('is case insensitive', () => {
      expect(Utils.getColumnIndex(headers, 'status')).toBe(3);
      expect(Utils.getColumnIndex(headers, 'STATUS')).toBe(3);
    });

    test('returns -1 for not found', () => {
      expect(Utils.getColumnIndex(headers, 'MISSING')).toBe(-1);
    });

    test('handles empty headers', () => {
      expect(Utils.getColumnIndex([], 'NAME')).toBe(-1);
    });
  });

  describe('getValue', () => {
    const headers = ['DATE', 'NAME', 'AMOUNT'];
    const row = ['2026-01-16', 'Test', 1000];

    test('gets value by column name', () => {
      expect(Utils.getValue(row, headers, 'NAME')).toBe('Test');
    });

    test('returns default for missing column', () => {
      expect(Utils.getValue(row, headers, 'MISSING', 'default')).toBe('default');
    });

    test('returns default for empty value', () => {
      const rowWithEmpty = ['2026-01-16', '', 1000];
      expect(Utils.getValue(rowWithEmpty, headers, 'NAME', 'default')).toBe('default');
    });
  });

  describe('parseDate', () => {
    test('returns Date objects as-is', () => {
      const date = new Date('2026-01-16');
      expect(Utils.parseDate(date)).toEqual(date);
    });

    test('parses date strings', () => {
      const result = Utils.parseDate('2026-01-16');
      expect(result instanceof Date).toBe(true);
      expect(result.getFullYear()).toBe(2026);
    });

    test('returns null for invalid dates', () => {
      expect(Utils.parseDate('not a date')).toBe(null);
      expect(Utils.parseDate('')).toBe(null);
      expect(Utils.parseDate(null)).toBe(null);
    });
  });

  describe('parseNumber', () => {
    test('returns numbers as-is', () => {
      expect(Utils.parseNumber(1000)).toBe(1000);
    });

    test('parses string numbers', () => {
      expect(Utils.parseNumber('1000')).toBe(1000);
      expect(Utils.parseNumber('1,000.50')).toBe(1000.50);
    });

    test('removes currency symbols', () => {
      expect(Utils.parseNumber('₹ 1,000')).toBe(1000);
      expect(Utils.parseNumber('$500')).toBe(500);
    });

    test('returns 0 for invalid numbers', () => {
      expect(Utils.parseNumber('')).toBe(0);
      expect(Utils.parseNumber(null)).toBe(0);
      expect(Utils.parseNumber('abc')).toBe(0);
    });
  });

  describe('formatDate', () => {
    const date = new Date('2026-01-16T10:30:00');

    test('formats with default format', () => {
      expect(Utils.formatDate(date)).toBe('16-01-2026');
    });

    test('formats with custom format', () => {
      expect(Utils.formatDate(date, 'YYYY-MM-DD')).toBe('2026-01-16');
    });

    test('returns empty string for invalid date', () => {
      expect(Utils.formatDate(null)).toBe('');
      expect(Utils.formatDate('not a date')).toBe('');
    });
  });

  describe('formatCurrency', () => {
    test('formats with default currency', () => {
      const result = Utils.formatCurrency(1000);
      expect(result).toContain('₹');
      expect(result).toContain('1,000.00');
    });

    test('formats with custom currency', () => {
      const result = Utils.formatCurrency(500, '$');
      expect(result).toContain('$');
    });

    test('handles null/undefined', () => {
      expect(Utils.formatCurrency(null)).toBe('');
      expect(Utils.formatCurrency(undefined)).toBe('');
    });
  });

  describe('chunk', () => {
    test('splits array into chunks', () => {
      const arr = [1, 2, 3, 4, 5];
      const chunks = Utils.chunk(arr, 2);
      expect(chunks).toEqual([[1, 2], [3, 4], [5]]);
    });

    test('handles empty array', () => {
      expect(Utils.chunk([], 2)).toEqual([]);
    });

    test('handles chunk size larger than array', () => {
      expect(Utils.chunk([1, 2], 5)).toEqual([[1, 2]]);
    });
  });

  describe('validateRequired', () => {
    test('returns valid for complete object', () => {
      const obj = { name: 'Test', amount: 100 };
      const result = Utils.validateRequired(obj, ['name', 'amount']);
      expect(result.valid).toBe(true);
      expect(result.missing).toEqual([]);
    });

    test('returns invalid for missing fields', () => {
      const obj = { name: 'Test' };
      const result = Utils.validateRequired(obj, ['name', 'amount']);
      expect(result.valid).toBe(false);
      expect(result.missing).toContain('amount');
    });

    test('treats empty string as missing', () => {
      const obj = { name: '' };
      const result = Utils.validateRequired(obj, ['name']);
      expect(result.valid).toBe(false);
    });
  });

});
