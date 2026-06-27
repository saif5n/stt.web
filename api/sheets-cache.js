const CACHE = {};

function now() { return Date.now(); }

async function batchGetCached(sheets, spreadsheetId, ranges, ttl = 30000) {
  const key = `batch:${ranges.join('|')}`;
  const entry = CACHE[key];
  if (entry && (now() - entry.ts) < ttl) {
    return entry.value;
  }
  const resp = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges, majorDimension: 'ROWS' });
  CACHE[key] = { ts: now(), value: resp };
  return resp;
}

function clearCache() { Object.keys(CACHE).forEach(k => delete CACHE[k]); }

module.exports = { batchGetCached, clearCache };
