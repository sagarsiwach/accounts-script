/**
 * UI Helper Functions
 * Server-side functions called by the sidebar
 *
 * @fileoverview UI support functions for CG Accounts
 */

/**
 * Gets active organizations for the sidebar dropdown
 * @returns {Array} Active orgs
 */
function getActiveOrgs() {
  const config = Config.load();
  return config.orgs.filter(org => org.active).map(org => ({
    code: org.code,
    name: org.name,
    active: org.active
  }));
}

/**
 * Gets recent logs for UI display
 * @returns {Array} Recent log entries
 */
function getRecentLogsForUI() {
  return AppLogger.getRecentLogs(10);
}

/**
 * Gets summary stats for the dashboard
 * @returns {Object} Stats object
 */
function getDashboardStats() {
  const config = Config.load();
  const logs = AppLogger.getRecentLogs(100);

  // Get last 24 hours of logs
  const cutoff = new Date();
  cutoff.setHours(cutoff.getHours() - 24);

  const recentLogs = logs.filter(log => new Date(log.timestamp) > cutoff);

  const successCount = recentLogs.filter(l => l.status === 'SUCCESS').length;
  const errorCount = recentLogs.filter(l => l.status === 'ERROR').length;
  const totalRows = recentLogs.reduce((sum, l) => sum + (l.rowsFetched || 0) + (l.rowsWritten || 0), 0);

  return {
    activeOrgs: config.orgs.filter(o => o.active).length,
    runsLast24h: recentLogs.length,
    successRate: recentLogs.length > 0 ? Math.round((successCount / recentLogs.length) * 100) : 0,
    errorCount: errorCount,
    totalRowsProcessed: totalRows
  };
}
