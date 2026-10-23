const NodeCache = require('node-cache');

// Step 2: Initialize NodeCache instance with standard TTL of 60 seconds
const cache = new NodeCache({
  stdTTL: 60,       // Default Time-To-Live: 60 seconds
  checkperiod: 120, // Automatic expired key cleanup every 120 seconds
});

// Cache performance metrics tracking (Supplementary Problem 2)
const cacheStats = {
  hits: 0,
  misses: 0,
  invalidations: 0,
};

/**
 * Custom get with automatic hit/miss instrumentation
 */
const getCache = (key) => {
  const value = cache.get(key);
  if (value !== undefined) {
    cacheStats.hits += 1;
    return value;
  }
  cacheStats.misses += 1;
  return null;
};

/**
 * Custom set with optional custom TTL
 */
const setCache = (key, value, ttl = 60) => {
  return cache.set(key, value, ttl);
};

/**
 * Invalidate task-related caches on write operations (Step 4 & Supplementary 1)
 */
const invalidateTaskCache = (taskId = null) => {
  cacheStats.invalidations += 1;

  // Invalidate all tasks collection and filtered queries
  const keys = cache.keys();
  keys.forEach((key) => {
    if (key.startsWith('all_tasks') || key.startsWith('tasks_')) {
      cache.del(key);
    }
  });

  // Invalidate specific single-task cache if ID is provided (Supplementary 1)
  if (taskId) {
    cache.del(`task_${taskId}`);
  }
};

/**
 * Retrieve aggregated cache diagnostics and statistics
 */
const getDiagnostics = () => {
  const total = cacheStats.hits + cacheStats.misses;
  const hitRate = total > 0 ? `${((cacheStats.hits / total) * 100).toFixed(2)}%` : '0.00%';
  const nodeStats = cache.getStats();

  return {
    hits: cacheStats.hits,
    misses: cacheStats.misses,
    totalRequests: total,
    hitRate,
    invalidations: cacheStats.invalidations,
    cachedKeys: cache.keys(),
    keyCount: cache.keys().length,
    stdTTL: 60,
    nodeCacheInternal: nodeStats,
  };
};

module.exports = {
  cache,
  getCache,
  setCache,
  invalidateTaskCache,
  getDiagnostics,
};
