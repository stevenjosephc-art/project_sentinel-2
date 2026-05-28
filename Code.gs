// The ONE Master Database ID
const MASTER_DB_ID = '1GO1-wYwIEIk5_2WQzgNvDF7gHYFxCmQA-azC_6QmzuU';

const SCHEMA = {
  'Open Sup Cases': [
    'Timestamp', 'Email', 'LDAP', 'Case ID', 'Symptom',
    'Detailed Issue', 'Reason', 'Channel', 'Team', 'Case ID Link',
    'Date', 'Time', 'Time Spent Before Taken', 'Handled By',
    'SME Remarks', 'Claimed At', 'Resolution Type',
    'Warn 24hr Sent', 'Warn 48hr Sent'
  ],
  'SupervisorList': [
    'Email Address', 'LDAP', 'Role'
    // cols D+ are dynamic date columns — not validated here
  ],
  'Metrics': [
    'Snapshot Time', 'Open Queue', 'Unclaimed', 'Claimed',
    'Resolved Today', 'Avg TAT Today (mins)',
    'Top Symptom Today', 'Top Channel Today',
    'Top SME Today', 'Top Escalation Driver Today'
  ],
  'Audit Log': [
    'Timestamp', 'Actor Email', 'Actor LDAP', 'Action',
    'Target Row', 'Case ID', 'Detail', 'IP / Session'
  ]
};

// ── RATE LIMITING ─────────────────────────────────────────────────
const RATE_LIMITS = {
  submitEscalation : { max: 5,  windowMins: 10  },
  claimCase        : { max: 20, windowMins: 5   },
  resolveCase      : { max: 10, windowMins: 10  },
  getOpenCases     : { max: 60, windowMins: 1   },
  getResolvedCases : { max: 30, windowMins: 1   },
  getSchedule      : { max: 30, windowMins: 1   },
};

function _checkRateLimit(action) {
  const cache    = CacheService.getUserCache();
  const key      = 'rl_' + action;
  const limit    = RATE_LIMITS[action];
  if (!limit) return; // no limit defined → allow

  const raw      = cache.get(key);
  const count    = raw ? parseInt(raw, 10) : 0;

  if (count >= limit.max) {
    _log('RATE_LIMIT', _actorEmail(), action + ' blocked (' + count + ' calls)');
    throw new Error('Rate limit exceeded for ' + action + '. Please wait a moment and try again.');
  }

  cache.put(key, String(count + 1), limit.windowMins * 60);
}

function _ensureSchema() {
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);

  Object.keys(SCHEMA).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      _log('SCHEMA', 'system', 'Created missing sheet: ' + sheetName);
    }

    const expectedHeaders = SCHEMA[sheetName];
    if (!expectedHeaders.length) return;

    const existingHeaders = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim())
      : [];

    let dirty = false;
    expectedHeaders.forEach((col, i) => {
      if (existingHeaders[i] !== col) {
        sheet.getRange(1, i + 1).setValue(col);
        dirty = true;
      }
    });

    if (dirty) {
      sheet.getRange(1, 1, 1, expectedHeaders.length)
        .setFontWeight('bold')
        .setBackground('#1A73E8')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      _log('SCHEMA', 'system', 'Repaired headers for sheet: ' + sheetName);
    }
  });
}

// SupervisorList sheet name
const SHEET_SUPERS = 'SupervisorList';

function getSessionAndRole() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  var ldap = email.split('@')[0].toLowerCase().trim();

  var myTeam = null;
  for (var team in TEAM_SUPERVISOR_MAP) {
    if (TEAM_SUPERVISOR_MAP[team] === ldap) { myTeam = team; break; }
  }
  if (!myTeam) {
    for (var team in TEAM_SME_MAP) {
      if ((TEAM_SME_MAP[team] || []).indexOf(ldap) > -1) { myTeam = team; break; }
    }
  }

  return {
    email: email || '',
    isSME: isUserSME_(email),
    myTeam: myTeam || null
  };
}

function isUserSME_(email) {
  if (!email) return false;
  try {
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet = ss.getSheetByName(SHEET_SUPERS);
    if (!sheet) return false;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const normalised = email.trim().toLowerCase();
    for (var i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).trim().toLowerCase() === normalised) return true;
    }
  } catch (e) {}
  return false;
}

function requireSME_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!isUserSME_(email)) {
    throw new Error('Access denied: only Supervisors and SMEs can perform this action.');
  }
}

function doGet(e) {
  _ensureSchema();
  if (e && e.parameter && e.parameter.page === 'schedule') {
    return HtmlService.createTemplateFromFile('Schedule')
        .evaluate().setTitle('Team Schedule - Google Play')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // Default to Dashboard
  return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate().setTitle('Escalations - Google Play')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSessionEmail() {
  try { return Session.getActiveUser().getEmail(); }
  catch (e) { return 'User Email Unavailable'; }
}

function submitEscalation(formData) {
  _checkRateLimit('submitEscalation');
  try {
    const submitterEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    let sheet = ss.getSheetByName('Open Sup Cases');

    if (!sheet) {
      sheet = ss.insertSheet('Open Sup Cases');
      const headers = ['Timestamp', 'Email', 'LDAP', 'Case ID', 'Symptom', 'Detailed Issue', 'Reason', 'Channel', 'Team'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF').setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    if (!/^\d-\d{13}$/.test(formData.caseId.trim())) return { success: false, message: 'Invalid Case ID format. Must be X-XXXXXXXXXXXXX (13 digits after dash).' };

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const allData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      const isDuplicate = allData.some(row =>
        String(row[2]).toLowerCase().trim() === formData.ldap.trim().toLowerCase() &&
        String(row[3]).trim()               === formData.caseId.trim()             &&
        String(row[4]).trim()               === formData.symptom.trim()
      );
      if (isDuplicate) return { success: false, message: 'Duplicate entry found.' };
    }

    const rowToAppend = [
      new Date(), submitterEmail, formData.ldap.trim(), formData.caseId.trim(),
      formData.symptom.trim(), formData.detailedIssue.trim(), formData.reason.trim(),
      formData.channel, formData.team
    ];

    sheet.appendRow(rowToAppend);
_log('SUBMIT', submitterEmail, 'Case submitted', -1, formData.caseId, formData.ldap + ' | ' + formData.symptom);

    return { success: true, message: 'Escalation submitted successfully!' };
  } catch (err) { return { success: false, message: 'Server Error: ' + err.message }; }
}

function pingCase(caseId) {
  const email = Session.getActiveUser().getEmail();
  if (!isUserSME_(email)) return;
  const ldap = email.split('@')[0].toLowerCase();

  const cache = CacheService.getScriptCache();
  const key = 'viewers_' + caseId;
  let viewers = JSON.parse(cache.get(key) || '[]');

  const now = Date.now();
  viewers = viewers.filter(v => v.ldap !== ldap && (now - v.ts) < 120000);
  viewers.push({ ldap: ldap, ts: now });

  cache.put(key, JSON.stringify(viewers), 125);
}

/**
 * Pings a case to indicate active viewing (SME Collision Detection).
 */
function pingCase(caseId) {
  const email = Session.getActiveUser().getEmail();
  if (!isUserSME_(email)) return;
  const ldap = email.split('@')[0].toLowerCase();

  const cache = CacheService.getScriptCache();
  const key = 'viewers_' + caseId;
  let viewers = JSON.parse(cache.get(key) || '[]');

  // Remove expired (older than 2 mins) or duplicate
  const now = Date.now();
  viewers = viewers.filter(v => v.ldap !== ldap && (now - v.ts) < 120000);
  viewers.push({ ldap: ldap, ts: now });

  cache.put(key, JSON.stringify(viewers), 125); // Cache for slightly more than 2 mins
}

function getOpenCases() {
  _checkRateLimit('getOpenCases');
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName('Open Sup Cases');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const openCases = [];
  const myEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  const myLdap = myEmail.split('@')[0];

  const INDEX_REMARKS = 14;
  const INDEX_HANDLED = 13;

  data.forEach((row, index) => {
    const remarks   = String(row[INDEX_REMARKS] || '').trim();
    if (remarks !== '') return;

    const handledBy = String(row[INDEX_HANDLED] || '').trim();
    const claimLdap = handledBy.toLowerCase().split('@')[0];
    const isMine    = (handledBy.toLowerCase() === myEmail) ||
                      (claimLdap === myLdap && claimLdap !== '');

    const viewers = JSON.parse(CacheService.getScriptCache().get('viewers_' + row[3]) || '[]');
    const now = Date.now();
    const activeViewers = viewers
      .filter(v => (now - v.ts) < 120000 && v.ldap !== myLdap)
      .map(v => v.ldap);

    openCases.push({
      rowIdx        : index + 2,
      viewers       : activeViewers,
      timestamp     : row[0]  ? new Date(row[0]).toString()  : '',
      submitter     : String(row[1]  || ''),
      ldap          : String(row[2]  || ''),
      caseId        : String(row[3]  || ''),
      symptom       : String(row[4]  || ''),
      detailedIssue : String(row[5]  || ''),
      reason        : String(row[6]  || ''),
      channel       : String(row[7]  || ''),
      team          : String(row[8]  || ''),
      caseLink      : String(row[9]  || ''),
      claimedAt     : row[15] ? new Date(row[15]).toString() : '',
      status        : handledBy !== '' ? 'In Progress' : 'Open',
      claimedBy     : handledBy,
      isMine        : isMine
    });
  });

  return openCases;
}

function getResolvedCases() {
  _checkRateLimit('getResolvedCases');
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName('Open Sup Cases');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rowLimit = Math.min(lastRow - 1, 3000);
  const startRow = Math.max(2, lastRow - rowLimit + 1);

  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 17).getValues();
  const resolvedCases = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const remarks = String(row[14] || '').trim();

    if (remarks !== "") {
      resolvedCases.push({
        rowIdx        : i + startRow,
        timestamp     : row[0] ? new Date(row[0]).toString() : '',
        submitter     : String(row[1] || ''),
        ldap          : String(row[2] || ''),
        caseId        : String(row[3] || ''),
        symptom       : String(row[4] || ''),
        detailedIssue : String(row[5] || ''),
        reason        : String(row[6] || ''),
        channel       : String(row[7] || ''),
        team          : String(row[8] || ''),
        caseLink      : String(row[9] || ''),
        claimedAt     : row[15] ? new Date(row[15]).toString() : '',
        resolutionTime: row[12] ? new Date(row[12]).toString() : (row[0] ? new Date(row[0]).toString() : ''),
        handledBy     : String(row[13] || ''),
        remarks       : remarks,
        resolutionType: String(row[16] || ''),
      });
    }
  }

  return resolvedCases;
}

function getSchedule() {
  _checkRateLimit('getSchedule');
  try {
    const ss    = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet = ss.getSheetByName(SHEET_SUPERS);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 4) return [];

    const headerRange = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

    const PHT_OFFSET_MS = 8 * 60 * 60 * 1000;
    const nowPHT        = new Date(Date.now() + PHT_OFFSET_MS);

    const targetDates = [0, 1, 2].map(function(offset) {
      const d = new Date(nowPHT);
      d.setUTCDate(d.getUTCDate() + offset);
      return d;
    });

    function makeDateLabel(d) {
      const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return months[d.getUTCMonth()] + '-' + d.getUTCDate();
    }

    const targetColIndices = targetDates.map(function(d) {
      const lbl = makeDateLabel(d);
      
      const idx = headerRange.findIndex(function(h) {
        if (h instanceof Date) {
          const hPHT = new Date(h.getTime() + PHT_OFFSET_MS);
          const hMonth = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][hPHT.getUTCMonth()];
          const hDate = hPHT.getUTCDate();
          return (hMonth + '-' + hDate) === lbl;
        }
        return String(h).trim() === lbl;
      });
      
      return { date: d, label: lbl, colIdx: idx };
    });

    const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    function makeDisplayLabel(d, offset) {
      if (offset === 0) return 'Today';
      if (offset === 1) return 'Tomorrow';
      return DAY_NAMES[d.getUTCDay()] + ' ' + makeDateLabel(d).replace('-', ' ');
    }

    const result = [];
    dataRange.forEach(function(row) {
      const email = String(row[0] || '').trim();
      const ldap  = String(row[1] || '').trim();
      const role  = String(row[2] || '').trim();
      if (!ldap) return;

      const days = targetColIndices.map(function(tc, offset) {
        var rawVal = '';
        if (tc.colIdx >= 0) {
          rawVal = String(row[tc.colIdx + 3] || '').trim();
        }

        var status = 'unknown';
        var displayValue = rawVal || '—';

        if (rawVal === '') {
          status = 'unknown';
        } else if (rawVal.toUpperCase() === 'OFF') {
          status = 'off';
          displayValue = 'Day Off';
        } else if (rawVal.toUpperCase() === 'VL') {
          status = 'vl';
          displayValue = 'On Leave';
        } else {
          status = 'scheduled';
          displayValue = rawVal;

          if (offset === 0) {
            function getShift(baseDate, timeStr) {
              var m = timeStr.trim().match(/^(\d+):(\d+):(\d+)\s+(AM|PM)$/i);
              if (!m) return null;
              var h = parseInt(m[1], 10);
              if (m[4].toUpperCase() === 'PM' && h < 12) h += 12;
              if (m[4].toUpperCase() === 'AM' && h === 12) h = 0;
              
              var start = new Date(baseDate.getTime());
              start.setUTCHours(h, parseInt(m[2], 10), 0, 0); 
              return { start: start, end: new Date(start.getTime() + (9 * 60 * 60 * 1000)) };
            }

            var shiftToday = getShift(tc.date, rawVal);
            var yestVal = tc.colIdx > 0 ? String(row[tc.colIdx + 3 - 1] || '').trim() : '';
            var yestDate = new Date(tc.date.getTime() - 86400000);
            var shiftYest = getShift(yestDate, yestVal);

            var isActive = false;
            if (shiftToday && nowPHT >= shiftToday.start && nowPHT < shiftToday.end) isActive = true;
            if (shiftYest && nowPHT >= shiftYest.start && nowPHT < shiftYest.end) isActive = true;

            if (isActive) status = 'on';

          } else {
            status = 'on';
          }
        }

        return {
          date        : tc.date.toISOString().slice(0, 10),
          label       : makeDisplayLabel(tc.date, offset),
          dateLabel   : tc.label,
          value       : displayValue,
          status      : status
        };
      });

      result.push({ email, ldap, role, days });
    });

    return result;
  } catch(e) {
    console.error('getSchedule error: ' + e.message);
    return [];
  }
}

function recordMetricsSnapshot() {
  try {
    const ss          = SpreadsheetApp.openById(MASTER_DB_ID);
    const openSheet   = ss.getSheetByName('Open Sup Cases');
    if (!openSheet) return;

    let metricsSheet = ss.getSheetByName('Metrics');
    if (!metricsSheet) {
      metricsSheet = ss.insertSheet('Metrics');
      const headers = [
        'Snapshot Time', 'Open Queue', 'Unclaimed', 'Claimed',
        'Resolved Today', 'Avg TAT Today (mins)',
        'Top Symptom Today', 'Top Channel Today', 'Top SME Today', 'Top Escalation Driver Today'
      ];
      metricsSheet.appendRow(headers);
      metricsSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1A73E8')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      metricsSheet.setFrozenRows(1);
    }

    const lastRow = openSheet.getLastRow();
    if (lastRow < 3) {
      metricsSheet.appendRow([new Date(), 0, 0, 0, 0, '', '—', '—', '—']);
      return;
    }

    const data = openSheet.getRange(2, 1, lastRow - 1, 17).getValues();

    const PHT_OFFSET_MS = 8 * 60 * 60 * 1000;
    const nowUTC        = Date.now();
    const nowPHT        = new Date(nowUTC + PHT_OFFSET_MS);
    const midnightPHT   = new Date(Date.UTC(
      nowPHT.getUTCFullYear(), nowPHT.getUTCMonth(), nowPHT.getUTCDate()
    ));

    let openCount    = 0;
    let claimedCount = 0;
    let resolvedToday = 0;
    let tatTotalMins  = 0;
    let tatCount      = 0;
    const symptoms   = {};
    const resTypes = {};
    const channels   = {};
    const smes       = {};

    data.forEach(function(row) {
      const handledBy    = String(row[13] || '').trim();
      const remarks      = String(row[14] || '').trim();
      const claimedAt    = row[15];
      const resolutionTs = row[12];

      if (remarks === '') {
        openCount++;
        if (handledBy !== '') claimedCount++;
      } else {
        if (resolutionTs) {
          const resDate = new Date(resolutionTs);
          if (!isNaN(resDate.getTime()) && resDate.getTime() >= midnightPHT.getTime()) {
            resolvedToday++;
            const startTs = claimedAt ? new Date(claimedAt) : new Date(row[0]);
            if (!isNaN(startTs.getTime())) {
              const diff = (resDate.getTime() - startTs.getTime()) / 60000;
              if (diff > 0) { tatTotalMins += diff; tatCount++; }
            }
            const sym = String(row[4] || '').trim();
            if (sym) symptoms[sym] = (symptoms[sym] || 0) + 1;
            const chan = String(row[7] || '').trim();
            if (chan) channels[chan] = (channels[chan] || 0) + 1;
            const smeLdap = handledBy ? handledBy.split('@')[0].toLowerCase() : '';
            if (smeLdap) smes[smeLdap] = (smes[smeLdap] || 0) + 1;
            const rType = String(row[16] || '').trim();
            if (rType) resTypes[rType] = (resTypes[rType] || 0) + 1;
          }
        }
      }
    });

    const avgTat     = tatCount > 0 ? Math.round(tatTotalMins / tatCount) : 0;
    const topSymptom = Object.keys(symptoms).sort(function(a,b){ return symptoms[b]-symptoms[a]; })[0] || '—';
    const topChannel = Object.keys(channels).sort(function(a,b){ return channels[b]-channels[a]; })[0] || '—';
    const topSME     = Object.keys(smes).sort(function(a,b){ return smes[b]-smes[a]; })[0] || '—';
    const topResType = Object.keys(resTypes).sort(function(a,b){ return resTypes[b]-resTypes[a]; })[0] || '—';

    metricsSheet.appendRow([
  new Date(), openCount, openCount - claimedCount, claimedCount,
  resolvedToday, avgTat > 0 ? avgTat : '—',
  topSymptom, topChannel, topSME, topResType
]);

  } catch(e) {
    console.error('recordMetricsSnapshot error: ' + e.message);
  }
}

function setupHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'recordMetricsSnapshot') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('recordMetricsSnapshot')
    .timeBased()
    .everyHours(1)
    .create();
}

function claimCase(rowIdx) {
  _checkRateLimit('claimCase');
  requireSME_();
  try {
    const ss        = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet     = ss.getSheetByName('Open Sup Cases');
    const smeEmail  = Session.getActiveUser().getEmail();
    const claimedAt = new Date();

    sheet.getRange(rowIdx, 16).setValue(claimedAt);
    sheet.getRange(rowIdx, 14).setValue(smeEmail);

    const row = sheet.getRange(rowIdx, 1, 1, 16).getValues()[0];
    const caseData = {
      submitter    : String(row[1] || ''),
      ldap         : String(row[2] || ''),
      caseId       : String(row[3] || ''),
      symptom      : String(row[4] || ''),
      detailedIssue: String(row[5] || ''),
      reason       : String(row[6] || ''),
      channel      : String(row[7] || ''),
      team         : String(row[8] || ''),
      caseLink     : String(row[9] || ''),
      timestamp    : row[0] ? new Date(row[0]).toString() : ''
    };

    _log('CLAIM', smeEmail, 'Case claimed', rowIdx, caseData.caseId, caseData.ldap);
    sendClaimNotification_(caseData, smeEmail);
    return { success: true, message: 'Case claimed.', email: smeEmail };
  } catch (e) { return { success: false, message: e.message }; }
}

function claimMultipleCases(rowIdxArray) {
  _checkRateLimit('claimCase');
  requireSME_();
  try {
    const ss        = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet     = ss.getSheetByName('Open Sup Cases');
    const smeEmail  = Session.getActiveUser().getEmail();
    const claimedAt = new Date();

    rowIdxArray.forEach(idx => {
      const currentClaim = sheet.getRange(idx, 14).getValue();
      if (!currentClaim) {
        sheet.getRange(idx, 16).setValue(claimedAt);
        sheet.getRange(idx, 14).setValue(smeEmail);

        const row = sheet.getRange(idx, 1, 1, 16).getValues()[0];
        const caseData = {
          submitter    : String(row[1] || ''),
          ldap         : String(row[2] || ''),
          caseId       : String(row[3] || ''),
          symptom      : String(row[4] || ''),
          detailedIssue: String(row[5] || ''),
          reason       : String(row[6] || ''),
          channel      : String(row[7] || ''),
          team         : String(row[8] || ''),
          caseLink     : String(row[9] || ''),
          timestamp    : row[0] ? new Date(row[0]).toString() : ''
        };
        sendClaimNotification_(caseData, smeEmail);
        _log('BULK_CLAIM', smeEmail, 'Bulk claim', idx, caseData.caseId, caseData.ldap);
      }
    });

    return { success: true, message: rowIdxArray.length + ' cases claimed.' };
  } catch (e) { return { success: false, message: e.message }; }
}

function resolveCase(rowIdx, remarks, resolutionType) {
  _checkRateLimit('resolveCase');
  requireSME_();
  try {
    const ss             = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet          = ss.getSheetByName('Open Sup Cases');
    const smeEmail       = Session.getActiveUser().getEmail();
    const resolutionTime = new Date();

    const row = sheet.getRange(rowIdx, 1, 1, 16).getValues()[0];
    const caseData = {
      submitter      : String(row[1] || ''),
      ldap           : String(row[2] || ''),
      caseId         : String(row[3] || ''),
      symptom        : String(row[4] || ''),
      detailedIssue  : String(row[5] || ''),
      reason         : String(row[6] || ''),
      channel        : String(row[7] || ''),
      team           : String(row[8] || ''),
      caseLink       : String(row[9] || ''),
      timestamp      : row[0] ? new Date(row[0]).toString() : '',
      resolutionTime : resolutionTime.toString(),
      remarks        : remarks
    };

    sheet.getRange(rowIdx, 13).setValue(resolutionTime);
    sheet.getRange(rowIdx, 15).setValue(remarks);
    sheet.getRange(rowIdx, 17).setValue(resolutionType || '');
    _log('RESOLVE', smeEmail, 'Case resolved | Driver: ' + resolutionType, rowIdx, caseData.caseId, remarks.substring(0, 80));

    sendResolveNotification_(caseData, smeEmail);
    hideRows();

    return { success: true, message: 'Case resolved successfully!' };
  } catch (e) { return { success: false, message: e.message }; }
}

function sendClaimNotification_(caseData, smeEmail) {
  try {
    if (!caseData.submitter) return;

    const smeLdap    = smeEmail.split('@')[0];
    const agentLdap  = caseData.ldap;
    const subject    = '✅ Your escalation has been picked up — Case ' + caseData.caseId;
    const claimedAt  = new Date().toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' });
    const submittedAt = caseData.timestamp ? new Date(caseData.timestamp).toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' }) : '—';

    const details = [
      ['Case ID', caseData.caseId, true],
      ['Symptom', caseData.symptom],
      ['Channel', caseData.channel],
      ['Team', caseData.team],
      ['Case Link', 'https://cases.connect.corp.google.com/' + caseData.caseId, true, 'Open Case'],
      ['Submitted', submittedAt],
      ['Picked up at', claimedAt],
      ['Handled by', smeLdap]
    ];

    const body = _getPremiumEmailHtml({
      type: 'blue',
      title: 'Case Picked Up',
      badgeIcon: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/check_circle/default/48px.svg',
      badgeText: 'PICKED UP',
      message: `Hi <b>${agentLdap}</b>, your escalation has been <span style="background-color:#fff7e0; color:#b93815; padding:2px 4px; border-radius:4px;">picked up</span> by <b>${smeLdap}</b> and is now <b style="color:#1a73e8">In Progress</b>.`,
      blocks: [
        { type: 'details', title: 'CASE DETAILS', icon: '📄', rows: details }
      ],
      footerNote: 'You will receive another notification once your case has been resolved. If you have urgent updates, please contact your supervisor directly.'
    });

    MailApp.sendEmail({
      to      : caseData.submitter,
      subject : subject,
      htmlBody: body,
      name    : 'Google Play Escalations',
      noReply : true,
      replyTo : 'play-escalations@google.com'
    });

  } catch (e) { console.error('sendClaimNotification_ error: ' + e.message); }
}

function sendResolveNotification_(caseData, smeEmail) {
  try {
    if (!caseData.submitter) return;

    const smeLdap     = smeEmail.split('@')[0];
    const agentLdap   = caseData.ldap;
    const subject     = '🎉 Your escalation has been resolved — Case ' + caseData.caseId;
    const resolvedAt  = new Date(caseData.resolutionTime).toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' });
    const submittedAt = caseData.timestamp ? new Date(caseData.timestamp).toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' }) : '—';

    const details = [
      ['Case ID', caseData.caseId],
      ['Symptom', caseData.symptom],
      ['Channel', caseData.channel],
      ['Team', caseData.team],
      ['Case Link', 'https://cases.connect.corp.google.com/' + caseData.caseId, true, 'Open Case'],
      ['Submitted', submittedAt],
      ['Resolved at', resolvedAt],
      ['Resolved by', smeLdap]
    ];

    const body = _getPremiumEmailHtml({
      type: 'green',
      title: 'Case Resolved',
      badgeIcon: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/celebration/default/48px.svg',
      badgeText: 'RESOLVED',
      message: `Hi <b>${agentLdap}</b>, your escalation has been <b style="color:#1e8e3e">resolved</b> by <b>${smeLdap}</b>.`,
      blocks: [
        { type: 'details', title: 'CASE DETAILS', icon: '📄', rows: details },
        { type: 'remarks', title: 'SME REMARKS', icon: '💬', content: caseData.remarks }
      ],
      footerNote: 'If you have any follow-up concerns, please contact your supervisor or submit a new escalation request.'
    });

    MailApp.sendEmail({
      to      : caseData.submitter,
      subject : subject,
      htmlBody: body,
      name    : 'Google Play Escalations',
      noReply : true,
      replyTo : 'play-escalations@google.com'
    });
  } catch (e) { console.error('sendResolveNotification_ error: ' + e.message); }
}

function escHtml_(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function _actorEmail() {
  try { return Session.getActiveUser().getEmail() || 'unknown'; }
  catch(e) { return 'unknown'; }
}

function _log(action, actor, detail, targetRow, caseId, extra) {
  try {
    const ss    = SpreadsheetApp.openById(MASTER_DB_ID);
    let logSheet = ss.getSheetByName('Audit Log');

    if (!logSheet) {
      logSheet = ss.insertSheet('Audit Log');
      logSheet.appendRow(['Timestamp','Actor Email','Actor LDAP','Action','Target Row','Case ID','Detail','Extra']);
      logSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
      logSheet.setFrozenRows(1);
    }

    const actorLdap = String(actor || '').split('@')[0];
    logSheet.appendRow([new Date(), actor || '', actorLdap || '', action || '', targetRow != null ? targetRow : '', caseId || '', detail || '', extra || '']);
  } catch(e) { console.error('_log failed: ' + e.message); }
}

function hideRows() {
  const ss    = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName("Open Sup Cases");
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const values = sheet.getRange(3, 15, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) { sheet.hideRows(3 + i); }
  }
}

function archiveOldCases() {
  const ss          = SpreadsheetApp.openById(MASTER_DB_ID);
  const sourceSheet = ss.getSheetByName('Open Sup Cases');
  let archiveSheet  = ss.getSheetByName('Archive');

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Archive');
    archiveSheet.appendRow(['Timestamp', 'Email', 'LDAP', 'Case ID', 'Symptom', 'Detailed Issue', 'Reason', 'Channel', 'Team', 'Case ID Link', 'Date', 'Time', 'Time spent before taken', 'Handled By', 'SME Remarks', 'Claimed At', 'Resolution Type']);
    archiveSheet.getRange(1, 1, 1, 17).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    archiveSheet.setFrozenRows(1);
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 3) return;

  const data          = sourceSheet.getRange(3, 1, lastRow - 2, 17).getValues();
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  let rowsToDelete = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const resolutionTime = data[i][12];
    if (resolutionTime && new Date(resolutionTime) < thirtyDaysAgo) {
      archiveSheet.appendRow(data[i]);
      rowsToDelete.push(i + 3);
    }
  }
  rowsToDelete.forEach(rowIdx => sourceSheet.deleteRow(rowIdx));
}

function getWebAppUrl() { return ScriptApp.getService().getUrl(); }

var TEAM_SUPERVISOR_MAP = {
  'Team Steven': 'stevenjosephc', 'Team Gerry': 'gerrymae', 'Team Khent': 'khent',
  'Team James': 'jamessevilla', 'Team Denden': 'bernardboy', 'Team Jim': 'jadoptante',
  'Team Al': 'acaluang', 'Team Faye': 'fajirnah', 'Team Mel': 'tapalla', 'Team Mary': 'marineth'
};

var TEAM_SME_MAP = {
  'Team Steven': ['sheenamae'], 'Team Gerry': ['caval'], 'Team Khent': ['criseldaa'],
  'Team James': ['acemile'], 'Team Denden': ['jgarlet'], 'Team Jim': ['glendajoe'],
  'Team Al': ['elagarto'], 'Team Faye': ['neljhon', 'mquirol'], 'Team Mel': ['fykeivan', 'mabubay'], 'Team Mary': ['fykeivan']
};

var SLA_CC_48HR = ['deanmark@google.com', 'joliveros@google.com', 'jmontrias@google.com'];

function checkSLAWarnings() {
  try {
    const ss        = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet     = ss.getSheetByName('Open Sup Cases');
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    const data    = sheet.getRange(3, 1, lastRow - 2, 19).getValues();
    const now     = new Date();
    const MS_24HR = 24 * 60 * 60 * 1000;
    const MS_48HR = 48 * 60 * 60 * 1000;

    data.forEach(function(row, i) {
      const remarks   = String(row[14] || '').trim();
      if (remarks !== '') return;

      const timestamp = row[0];
      if (!timestamp) return;

      const submitted = new Date(timestamp);
      if (isNaN(submitted.getTime())) return;

      const ageMs       = now.getTime() - submitted.getTime();
      const warn24Sent  = row[17] === true || String(row[17]).toUpperCase() === 'TRUE';
      const warn48Sent  = row[18] === true || String(row[18]).toUpperCase() === 'TRUE';

      const ldap        = String(row[2]    || '').trim();
      const caseId      = String(row[3]  || '').trim();
      const team        = String(row[8]    || '').trim();
      const sheetRow    = i + 3;

      const agentEmail  = ldap + '@google.com';
      const supLdap     = TEAM_SUPERVISOR_MAP[team] || '';
      const supEmail    = supLdap ? supLdap + '@google.com' : '';
      const smeLdaps    = TEAM_SME_MAP[team] || [];
      const smeEmails   = smeLdaps.map(function(s) { return s + '@google.com'; });
      const submittedStr = submitted.toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' });
      const hoursOld    = Math.floor(ageMs / (1000 * 60 * 60));

      if (ageMs >= MS_48HR && !warn48Sent) {
        sendSLA48hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, '', '', '', team, submittedStr, hoursOld);
        sheet.getRange(sheetRow, 18).setValue(true);
        sheet.getRange(sheetRow, 19).setValue(true);
        _log('SLA_48HR', 'system', 'SLA 48hr warning sent', sheetRow, caseId, ldap);
        return;
      }

      if (ageMs >= MS_24HR && !warn24Sent) {
        sendSLA24hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, '', '', '', team, submittedStr, hoursOld);
        sheet.getRange(sheetRow, 18).setValue(true);
        _log('SLA_24HR', 'system', 'SLA 24hr warning sent', sheetRow, caseId, ldap);
      }
    });
  } catch(e) { console.error('checkSLAWarnings error: ' + e.message); }
}

function sendSLA24hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, symptom, reason, channel, team, submittedStr, hoursOld, testMode_) {
  try {
    const subject = '⏰ Escalation Pending for 24 Hours — ' + caseId;
    const details = [
      ['Case ID', caseId, true],
      ['Agent', ldap],
      ['Symptom', symptom || 'Not specified'],
      ['Channel', channel || 'Not specified'],
      ['Team', team],
      ['Submitted', submittedStr],
      ['Age', hoursOld + ' hours in queue']
    ];

    const body = _getPremiumEmailHtml({
      type: 'orange',
      title: '24-Hour Pending Reminder',
      subtitle: 'Your escalation is still awaiting supervisor action.',
      badgeIcon: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/schedule/default/48px.svg',
      badgeText: '24h OLD',
      message: `Hi <b>${ldap}</b>, this is a friendly reminder that your escalation submitted <b>${submittedStr}</b> has been in the queue for over <b>24 hours</b> and is still awaiting action from a supervisor or SME.`,
      blocks: [
        { type: 'details', title: 'CASE DETAILS', icon: '📋', rows: details },
        {
          type: 'action',
          title: 'WHAT HAPPENS NEXT?',
          icon: '💡',
          color: 'blue',
          content: `A supervisor or SME from <b>${team}</b> will <span style="background-color:#fff7e0; color:#b93815; padding:2px 4px; border-radius:4px;">pick up</span> your case shortly. If this is urgent, please reach out to your team lead directly. You will receive another notification if your case remains unresolved at the 48-hour mark.`
        }
      ]
    });

    var ccList = [];
    if (!testMode_) {
      if (supEmail) ccList.push(supEmail);
      smeEmails.forEach(function(e) { if (e) ccList.push(e); });
    }

    MailApp.sendEmail({ to: agentEmail, cc: ccList.filter(Boolean).join(','), subject: subject, htmlBody: body, name: 'Google Play Escalations', noReply: true, replyTo: 'play-escalations@google.com' });
  } catch(e) { console.error('sendSLA24hrEmail_ error: ' + e.message); }
}

function sendSLA48hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, symptom, reason, channel, team, submittedStr, hoursOld, testMode_) {
  try {
    const subject = '🚨 URGENT: 48-Hour SLA Breach Warning — ' + caseId;
    const details = [
      ['Case ID', caseId, true],
      ['Agent', ldap],
      ['Symptom', symptom || 'Not specified'],
      ['Reason', reason || 'Not specified'],
      ['Channel', channel || 'Not specified'],
      ['Team', team],
      ['Submitted', submittedStr],
      ['Age', `<span style="color:#b3261e; font-weight:700;">${hoursOld} hours — SLA BREACHED</span>`]
    ];

    const body = _getPremiumEmailHtml({
      type: 'red',
      isUrgent: true,
      title: '48-Hour SLA Breach Warning',
      subtitle: 'This case has exceeded the 48-hour response threshold and requires immediate attention.',
      badgeIcon: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/notifications_active/default/48px.svg',
      badgeText: '48h OLD',
      message: `Hi <b>${ldap}</b>, your escalation submitted on <b>${submittedStr}</b> has now been pending for over <b>48 hours</b>. This exceeds the acceptable response threshold. Your team managers have been notified and this case requires <b style="color:#b3261e">immediate action</b>.`,
      alertBanner: `⚠️ This case has been in the queue for 48 hours without resolution. Management has been notified.`,
      blocks: [
        { type: 'details', title: 'CASE DETAILS', icon: '🚨', rows: details },
        {
          type: 'action',
          title: 'REQUIRED ACTION',
          icon: '⚡',
          color: 'yellow',
          content: `<ul style="margin:0; padding-left:20px;">
            <li>An SME or Supervisor from <b>${team}</b> must claim and resolve this case immediately.</li>
            <li>If the assigned team is unavailable, please escalate to your manager directly.</li>
            <li>Update the case in the <a href="${getWebAppUrl()}" style="color:#0b57d0; text-decoration:underline;">Escalations Dashboard</a> once resolved.</li>
          </ul>`
        },
        {
          type: 'action',
          title: 'MANAGERS NOTIFIED',
          icon: '✅',
          color: 'green',
          content: 'This alert has been automatically copied to the Play Ops management team for visibility and follow-up.'
        }
      ]
    });

    var ccList = [];
    if (supEmail) ccList.push(supEmail);
    smeEmails.forEach(function(e) { if (e) ccList.push(e); });
    if (!testMode_) { SLA_CC_48HR.forEach(function(e) { ccList.push(e); }); }

    MailApp.sendEmail({ to: agentEmail, cc: ccList.filter(Boolean).join(','), subject: subject, htmlBody: body, name: 'Google Play Escalations', noReply: true, replyTo: 'play-escalations@google.com' });
  } catch(e) { console.error('sendSLA48hrEmail_ error: ' + e.message); }
}

function _getPremiumEmailHtml(config) {
  const styles = {
    red:    { head: '#d93025', bg: '#fce8e6', text: '#b3261e', badge: '#c5221f' },
    orange: { head: '#f29900', bg: '#fff7e0', text: '#b93815', badge: '#e37400' },
    blue:   { head: '#1a73e8', bg: '#e8f0fe', text: '#1967d2', badge: '#185abc' },
    green:  { head: '#1e8e3e', bg: '#e6f4ea', text: '#137333', badge: '#137333' },
    yellow: { head: '#f9ab00', bg: '#fff7e0', text: '#b93815', badge: '#f29900' }
  };
  const theme = styles[config.type] || styles.blue;
  const headerSub = config.isUrgent ? 'GOOGLE PLAY · ESCALATIONS · URGENT ALERT' : 'GOOGLE PLAY · ESCALATIONS';

  let blocksHtml = '';
  (config.blocks || []).forEach(block => {
    if (block.type === 'details') {
      let rows = '';
      block.rows.forEach(r => {
        let val = r[1];
        if (r[2]) { // isLink/isCaseId
          if (r[3]) val = `<a href="${r[1]}" style="color:#0b57d0; text-decoration:underline; font-weight:700;">${r[3]}</a>`;
          else val = `<a href="https://cases.connect.corp.google.com/${r[1]}" style="color:#b3261e; text-decoration:underline; font-weight:700;">${r[1]}</a>`;
        }
        rows += `
          <tr style="border-bottom: 1px solid #f1f3f4;">
            <td style="padding: 10px 0; color: #5f6368; font-size: 13px; width: 140px; vertical-align: top;">${r[0]}</td>
            <td style="padding: 10px 0; color: #202124; font-size: 13px; font-weight: 500; vertical-align: top;">${val}</td>
          </tr>`;
      });
      blocksHtml += `
        <div style="background-color: ${theme.bg}; border-radius: 12px; padding: 24px; border-left: 6px solid ${theme.head}; margin-bottom: 24px; border: 1px solid #dadce0; border-left-width: 6px;">
          <div style="font-family: 'Google Sans', sans-serif; font-size: 11px; font-weight: 700; color: ${theme.text}; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 16px;">
            <span style="margin-right: 8px;">${block.icon || '📄'}</span> ${block.title}
          </div>
          <table style="width: 100%; border-collapse: collapse;">${rows}</table>
        </div>`;
    } else if (block.type === 'remarks') {
      blocksHtml += `
        <div style="background-color: #e6f4ea; border-radius: 12px; padding: 24px; border-left: 6px solid #1e8e3e; margin-bottom: 24px; border: 1px solid #c4eed0;">
          <div style="font-family: 'Google Sans', sans-serif; font-size: 11px; font-weight: 700; color: #0d652d; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 12px;">
            <span style="margin-right: 8px;">${block.icon || '💬'}</span> ${block.title}
          </div>
          <div style="font-size: 14px; line-height: 1.6; color: #072711; white-space: pre-wrap;">${escHtml_(block.content)}</div>
        </div>`;
    } else if (block.type === 'action') {
      const aTheme = styles[block.color] || theme;
      blocksHtml += `
        <div style="background-color: ${aTheme.bg}; border-radius: 12px; padding: 24px; border-left: 6px solid ${aTheme.head}; margin-bottom: 24px; border: 1px solid #dadce0; border-left-width: 6px;">
          <div style="font-family: 'Google Sans', sans-serif; font-size: 11px; font-weight: 700; color: ${aTheme.text}; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 12px;">
            <span style="margin-right: 8px;">${block.icon || '💡'}</span> ${block.title}
          </div>
          <div style="font-size: 14px; line-height: 1.6; color: #3c4043;">${block.content}</div>
        </div>`;
    }
  });

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500;700&display=swap');
        body { margin: 0; padding: 0; background-color: #f8f9fa; font-family: 'Roboto', Arial, sans-serif; }
      </style>
    </head>
    <body style="margin: 0; padding: 20px; background-color: #f8f9fa;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.08); border: 1px solid #e0e0e0;">
        <!-- Header -->
        <div style="background: ${theme.head}; padding: 32px; color: #ffffff; position: relative;">
          <div style="display: table; width: 100%;">
            <div style="display: table-cell; vertical-align: top;">
              <img src="https://www.gstatic.com/images/branding/product/2x/google_play_64dp.png" width="40" height="40" style="margin-bottom: 16px;">
              <div style="font-family: 'Google Sans', sans-serif; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px; opacity: 0.9;">${headerSub}</div>
              <h1 style="font-family: 'Google Sans', sans-serif; font-size: 28px; font-weight: 700; margin: 0; line-height: 1.2;">${config.title}</h1>
              ${config.subtitle ? `<div style="font-size: 14px; margin-top: 8px; opacity: 0.9;">${config.subtitle}</div>` : ''}
            </div>
            <div style="display: table-cell; vertical-align: middle; width: 100px; text-align: right;">
              <div style="background: rgba(255,255,255,0.2); border: 1px solid rgba(255,255,255,0.3); border-radius: 12px; padding: 12px; text-align: center;">
                <img src="${config.badgeIcon}" width="32" height="32" style="filter: brightness(0) invert(1); margin-bottom: 4px;">
                <div style="font-size: 10px; font-weight: 700; font-family: 'Google Sans', sans-serif;">${config.badgeText}</div>
              </div>
            </div>
          </div>
        </div>

        ${config.alertBanner ? `
          <div style="background-color: #fce8e6; padding: 16px; text-align: center; color: #b3261e; font-size: 14px; font-weight: 700; border-bottom: 1px solid #f8d7da;">
            ${config.alertBanner}
          </div>
        ` : ''}

        <!-- Content -->
        <div style="padding: 32px;">
          <div style="font-size: 16px; line-height: 1.6; color: #3c4043; margin-bottom: 32px;">
            ${config.message}
          </div>

          ${blocksHtml}

          ${config.footerNote ? `<div style="font-size: 14px; color: #5f6368; line-height: 1.6; border-top: 1px solid #f1f3f4; padding-top: 24px;">${config.footerNote}</div>` : ''}
        </div>

        <!-- Brand Footer -->
        <div style="background-color: #f1f3f4; padding: 24px 32px; border-top: 1px solid #e0e0e0; display: flex; align-items: center; justify-content: space-between;">
           <div style="display: flex; align-items: center; font-size: 12px; color: #70757a;">
             <img src="https://www.gstatic.com/images/branding/product/2x/google_play_64dp.png" width="20" height="20" style="margin-right: 8px; opacity: 0.6;">
             Google Play Escalations - Automated Notification - Do not reply
           </div>
           ${config.isUrgent ? `<div style="color: #b3261e; font-size: 11px; font-weight: 700; font-family: 'Google Sans', sans-serif;">SLA BREACH - 48h</div>` : ''}
        </div>
      </div>
      <div style="text-align: center; margin-top: 24px; font-size: 11px; color: #9aa0a6;">
        This is an automated message from the Google Play Escalations Dashboard.
      </div>
    </body>
    </html>
  `;
}

function setupSLATrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkSLAWarnings') { ScriptApp.deleteTrigger(trigger); }
  });
  ScriptApp.newTrigger('checkSLAWarnings').timeBased().everyHours(1).create();
}
