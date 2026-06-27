const CACHE = {
  uids: { ts: 0, data: [] },
  urls: { ts: 0, data: [] },
  submissions: { ts: 0, data: [] }
};

const TTL = {
  uids: 1000 * 60 * 5, // 5 minutes
  urls: 1000 * 30,     // 30 seconds
  submissions: 1000 * 30
};

async function getUIDs(sheets, spreadsheetId) {
  const now = Date.now();
  if (now - CACHE.uids.ts < TTL.uids && CACHE.uids.data.length) return CACHE.uids.data;
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'UIDs!A2:B' });
  CACHE.uids.ts = now;
  CACHE.uids.data = resp.data.values || [];
  return CACHE.uids.data;
}

async function getURLs(sheets, spreadsheetId) {
  const now = Date.now();
  if (now - CACHE.urls.ts < TTL.urls && CACHE.urls.data.length) return CACHE.urls.data;
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'URLs!A2:E' });
  CACHE.urls.ts = now;
  CACHE.urls.data = resp.data.values || [];
  return CACHE.urls.data;
}

async function getSubmissionsRange(sheets, spreadsheetId, range) {
  // range should be either 'Submissions!B:C' or 'Submissions!H:H' etc.
  const now = Date.now();
  // Simple keying: use submissions cache TTL for any submissions reads
  if (now - CACHE.submissions.ts < TTL.submissions && CACHE.submissions.dataMap && CACHE.submissions.dataMap[range]) {
    return CACHE.submissions.dataMap[range];
  }
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  CACHE.submissions.ts = now;
  CACHE.submissions.dataMap = CACHE.submissions.dataMap || {};
  CACHE.submissions.dataMap[range] = resp.data.values || [];
  return CACHE.submissions.dataMap[range];
}

function invalidateUIDs() { CACHE.uids.ts = 0; CACHE.uids.data = []; }
function invalidateURLs() { CACHE.urls.ts = 0; CACHE.urls.data = []; }
function invalidateSubmissions() { CACHE.submissions.ts = 0; CACHE.submissions.dataMap = {}; }

module.exports = {
  getUIDs,
  getURLs,
  getSubmissionsRange,
  invalidateUIDs,
  invalidateURLs,
  invalidateSubmissions
};
