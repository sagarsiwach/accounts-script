/**
 * Logger Module (LEGACY - DEPRECATED)
 *
 * @deprecated This module is deprecated. Use Init.logRun() instead.
 * This was designed for a centralized control sheet architecture.
 * Current architecture (v0.4.0+) uses per-sheet RUN_LOG tabs managed by init.js
 *
 * @fileoverview Legacy logging system - DO NOT USE
 */

const AppLogger = (function() {

  const LOG_SHEET_NAME = 'RUN_LOG';

  /**
   * Logs a completed run
   * @param {string} orgCode - Organization code
   * @param {Object} result - Run result object
   * @param {Date} startTime - When the run started
   */
  function logRun(orgCode, result, startTime) {
    const endTime = new Date();
    const duration = endTime - startTime;

    // Log each source separately
    for (const [sourceType, sourceResult] of Object.entries(result)) {
      if (sourceType === 'ledger') continue; // Log ledger separately

      writeLogEntry({
        timestamp: endTime,
        orgCode: orgCode,
        sourceType: sourceType.toUpperCase(),
        rowsFetched: sourceResult.rows || 0,
        rowsWritten: 0,
        status: sourceResult.status || 'SUCCESS',
        errorMessage: sourceResult.error || '',
        durationMs: Math.round(duration / 4) // Approximate per-source
      });
    }

    // Log ledger generation
    if (result.ledger) {
      writeLogEntry({
        timestamp: endTime,
        orgCode: orgCode,
        sourceType: 'LEDGER',
        rowsFetched: 0,
        rowsWritten: result.ledger.rows || 0,
        status: result.ledger.status || 'SUCCESS',
        errorMessage: result.ledger.error || '',
        durationMs: 0
      });
    }
  }

  /**
   * Logs an error
   * @param {string} orgCode - Organization code
   * @param {string} operation - What operation failed
   * @param {Error} error - The error object
   */
  function logError(orgCode, operation, error) {
    writeLogEntry({
      timestamp: new Date(),
      orgCode: orgCode,
      sourceType: operation,
      rowsFetched: 0,
      rowsWritten: 0,
      status: 'ERROR',
      errorMessage: error.message || String(error),
      durationMs: 0
    });
  }

  /**
   * Writes a log entry to the RUN_LOG sheet
   * @param {Object} entry - Log entry object
   */
  function writeLogEntry(entry) {
    try {
      const config = Config.load();
      const ss = SpreadsheetApp.openById(Config.CONTROL_SHEET_ID);
      let sheet = ss.getSheetByName(LOG_SHEET_NAME);

      // Create sheet if doesn't exist
      if (!sheet) {
        sheet = ss.insertSheet(LOG_SHEET_NAME);
        sheet.appendRow([
          'TIMESTAMP', 'ORG_CODE', 'SOURCE_TYPE', 'ROWS_FETCHED',
          'ROWS_WRITTEN', 'STATUS', 'ERROR_MESSAGE', 'DURATION_MS'
        ]);
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        entry.timestamp,
        entry.orgCode,
        entry.sourceType,
        entry.rowsFetched,
        entry.rowsWritten,
        entry.status,
        entry.errorMessage,
        entry.durationMs
      ]);

    } catch (e) {
      // If logging fails, at least log to Apps Script logs
      Logger.log('Failed to write log entry: ' + e.message);
      Logger.log('Original entry: ' + JSON.stringify(entry));
    }
  }

  /**
   * Sends error notification email
   * @param {string} subject - Email subject
   * @param {Array|Object} errors - Error details
   */
  function sendErrorEmail(subject, errors) {
    try {
      const config = Config.load();
      const recipients = config.settings.errorEmailRecipients;

      if (!recipients || recipients.length === 0) {
        Logger.log('No error email recipients configured');
        return;
      }

      const errorList = Array.isArray(errors) ? errors : [errors];

      let body = 'CG Accounts Script - Error Report\n\n';
      body += 'Time: ' + new Date().toISOString() + '\n\n';
      body += 'Errors:\n';
      body += '─'.repeat(50) + '\n';

      for (const error of errorList) {
        if (typeof error === 'object') {
          body += 'Org: ' + (error.org || 'N/A') + '\n';
          body += 'Error: ' + (error.error || error.message || JSON.stringify(error)) + '\n';
          body += '─'.repeat(50) + '\n';
        } else {
          body += error + '\n';
          body += '─'.repeat(50) + '\n';
        }
      }

      body += '\nPlease check the RUN_LOG sheet for more details.';

      MailApp.sendEmail({
        to: Array.isArray(recipients) ? recipients.join(',') : recipients,
        subject: '[CG Accounts] ' + subject,
        body: body
      });

    } catch (e) {
      Logger.log('Failed to send error email: ' + e.message);
    }
  }

  /**
   * Gets recent log entries
   * @param {number} limit - Max entries to return
   * @param {string} [orgCode] - Filter by org (optional)
   * @returns {Array} Log entries
   */
  function getRecentLogs(limit = 50, orgCode = null) {
    try {
      const ss = SpreadsheetApp.openById(Config.CONTROL_SHEET_ID);
      const sheet = ss.getSheetByName(LOG_SHEET_NAME);

      if (!sheet) return [];

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const logs = [];

      // Start from bottom (newest first)
      for (let i = data.length - 1; i > 0 && logs.length < limit; i--) {
        const row = data[i];

        // Filter by org if specified
        if (orgCode && row[headers.indexOf('ORG_CODE')] !== orgCode) {
          continue;
        }

        logs.push({
          timestamp: row[headers.indexOf('TIMESTAMP')],
          orgCode: row[headers.indexOf('ORG_CODE')],
          sourceType: row[headers.indexOf('SOURCE_TYPE')],
          rowsFetched: row[headers.indexOf('ROWS_FETCHED')],
          rowsWritten: row[headers.indexOf('ROWS_WRITTEN')],
          status: row[headers.indexOf('STATUS')],
          errorMessage: row[headers.indexOf('ERROR_MESSAGE')],
          durationMs: row[headers.indexOf('DURATION_MS')]
        });
      }

      return logs;

    } catch (e) {
      Logger.log('Failed to get logs: ' + e.message);
      return [];
    }
  }

  /**
   * Cleans up old log entries
   * @param {number} retentionDays - Days to keep
   * @returns {number} Rows deleted
   */
  function cleanupOldLogs(retentionDays = 30) {
    try {
      const ss = SpreadsheetApp.openById(Config.CONTROL_SHEET_ID);
      const sheet = ss.getSheetByName(LOG_SHEET_NAME);

      if (!sheet) return 0;

      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

      const data = sheet.getDataRange().getValues();
      const rowsToDelete = [];

      for (let i = 1; i < data.length; i++) {
        const timestamp = data[i][0];
        if (timestamp instanceof Date && timestamp < cutoffDate) {
          rowsToDelete.push(i + 1); // 1-indexed
        }
      }

      // Delete from bottom to top to preserve row indices
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
      }

      return rowsToDelete.length;

    } catch (e) {
      Logger.log('Failed to cleanup logs: ' + e.message);
      return 0;
    }
  }

  // Public API
  return {
    logRun,
    logError,
    sendErrorEmail,
    getRecentLogs,
    cleanupOldLogs
  };

})();
