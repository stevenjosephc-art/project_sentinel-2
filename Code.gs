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

    const detailsHtml =
      _detailRow('Case ID', caseData.caseId) +
      _detailRow('Symptom', caseData.symptom) +
      _detailRow('Channel', caseData.channel) +
      _detailRow('Team', caseData.team) +
      _detailRow('Case Link', 'https://cases.connect.corp.google.com/' + caseData.caseId, true) +
      _detailRow('Submitted', submittedAt) +
      _detailRow('Picked up at', claimedAt) +
      _detailRow('Handled by', smeLdap);

    const body = _getMaterial3EmailHtml({
      title: 'Case Picked Up',
      message: `Hi <b>${agentLdap}</b>, your escalation has been picked up by <b>${smeLdap}</b> and is now <b style="color:#0b57d0">In Progress</b>.`,
      detailsHtml: detailsHtml,
      buttonUrl: 'https://cases.connect.corp.google.com/' + caseData.caseId,
      buttonText: 'View Case',
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

    const detailsHtml =
      `<div style="background-color: #e6f4ea; border-radius: 8px; padding: 16px; margin-bottom: 24px; border: 1px solid #c4eed0;">
        <div style="font-size: 11px; font-weight: 700; color: #072711; text-transform: uppercase; margin-bottom: 8px;">Resolution Remarks</div>
        <div style="font-size: 14px; line-height: 1.5; color: #072711; white-space: pre-wrap;">${escHtml_(caseData.remarks)}</div>
      </div>` +
      _detailRow('Case ID', caseData.caseId) +
      _detailRow('Symptom', caseData.symptom) +
      _detailRow('Case Link', 'https://cases.connect.corp.google.com/' + caseData.caseId, true) +
      _detailRow('Submitted', submittedAt) +
      _detailRow('Resolved at', resolvedAt) +
      _detailRow('Handled by', smeLdap);

    const body = _getMaterial3EmailHtml({
      title: 'Case Resolved',
      headerBg: '#1e8e3e',
      message: `Hi <b>${agentLdap}</b>, your escalation regarding Case <b>${caseData.caseId}</b> has been <b style="color:#1e8e3e">Resolved</b>.`,
      detailsHtml: detailsHtml,
      buttonUrl: 'https://cases.connect.corp.google.com/' + caseData.caseId,
      buttonText: 'View Details',
      iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/celebration/default/48px.svg'
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
    const detailsHtml = _detailRow('Case ID', caseId) + _detailRow('Submitted', submittedStr) + _detailRow('Team', team) + _detailRow('Case Link', 'https://cases.connect.corp.google.com/' + caseId, true) + _detailRow('Waiting', `<span style="color:#b93815">${hoursOld} hours</span>`);

    const body = _getMaterial3EmailHtml({
      title: 'Still in Queue', headerBg: '#f29900', message: `Hi <b>${ldap}</b>, this is a friendly automated reminder that your escalation for Case <b>${caseId}</b> has been pending for over 24 hours.`, detailsHtml: detailsHtml, buttonUrl: 'https://cases.connect.corp.google.com/' + caseId, buttonText: 'View Queue', iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/schedule/default/48px.svg'
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
    const subject = '🚨 URGENT: SLA Breach for Case ' + caseId;
    const detailsHtml = _detailRow('Case ID', caseId) + _detailRow('Submitted', submittedStr) + _detailRow('Team', team) + _detailRow('Case Link', 'https://cases.connect.corp.google.com/' + caseId, true) + _detailRow('Queue Time', `<span style="color:#b3261e">${hoursOld} hours</span>`);

    const body = _getMaterial3EmailHtml({
      title: 'SLA Breach Alert', headerBg: '#b3261e', message: `Hi <b>${ldap}</b>, <b>Attention Required:</b> Escalation for Case <b>${caseId}</b> has exceeded the 48-hour threshold. Management has been notified.`, detailsHtml: detailsHtml, buttonUrl: 'https://cases.connect.corp.google.com/' + caseId, buttonText: 'Prioritize Case', iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/error/default/48px.svg'
    });

    var ccList = [];
    if (supEmail) ccList.push(supEmail);
    smeEmails.forEach(function(e) { if (e) ccList.push(e); });
    if (!testMode_) { SLA_CC_48HR.forEach(function(e) { ccList.push(e); }); }

    MailApp.sendEmail({ to: agentEmail, cc: ccList.filter(Boolean).join(','), subject: subject, htmlBody: body, name: 'Google Play Escalations', noReply: true, replyTo: 'play-escalations@google.com' });
  } catch(e) { console.error('sendSLA48hrEmail_ error: ' + e.message); }
}

function _getMaterial3EmailHtml(config) {
  const headerBg = config.headerBg || '#0b57d0';
  const iconUrl  = config.iconUrl || 'https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/check_circle/default/48px.svg';
  const buttonBg = config.buttonBg || headerBg;

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500;700&display=swap');
        body { margin: 0; padding: 0; background-color: #f1f3f4; font-family: 'Roboto', Arial, sans-serif; -webkit-font-smoothing: antialiased; }
        .wrapper { padding: 40px 20px; }
        .container { max-width: 640px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 1px 3px rgba(60,64,67,0.3), 0 4px 8px 3px rgba(60,64,67,0.15); border: 1px solid #dadce0; }
        .header { background-color: ${headerBg}; padding: 40px; color: #ffffff; position: relative; }
        .header-top { font-family: 'Google Sans', sans-serif; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; opacity: 0.95; }
        .header-title { font-family: 'Google Sans', sans-serif; font-size: 32px; font-weight: 500; margin: 0; }
        .header-icon { position: absolute; right: 40px; top: 50%; transform: translateY(-50%); background: rgba(255,255,255,0.25); border-radius: 12px; padding: 12px; display: flex; align-items: center; justify-content: center; }
        .content { padding: 40px; color: #3c4043; }
        .message { font-size: 16px; line-height: 1.6; margin-bottom: 32px; color: #3c4043; }
        .details-box { background-color: #f8f9fa; border-radius: 8px; padding: 32px; border-left: 4px solid ${headerBg}; margin-bottom: 32px; border: 1px solid #dadce0; }
        .details-header { font-family: 'Google Sans', sans-serif; font-size: 11px; font-weight: 700; color: #5f6368; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 24px; }
        .detail-row { display: table; width: 100%; margin-bottom: 16px; }
        .detail-label { display: table-cell; width: 200px; font-size: 13px; color: #5f6368; vertical-align: top; }
        .detail-value { display: table-cell; font-size: 13px; color: #202124; font-weight: 500; vertical-align: top; }
        .detail-link { color: #0b57d0; text-decoration: none; font-weight: 500; border-bottom: 1px solid #0b57d0; }
        .button-container { text-align: center; margin: 32px 0; }
        .button { background-color: ${buttonBg}; color: #ffffff !important; padding: 12px 32px; border-radius: 100px; text-decoration: none; font-family: 'Google Sans', sans-serif; font-size: 14px; font-weight: 500; display: inline-block; }
        .footer-note { font-size: 14px; color: #5f6368; line-height: 1.6; margin-top: 32px; }
        .brand-footer { padding: 32px 40px; background-color: #ffffff; border-top: 1px solid #f1f3f4; font-size: 12px; color: #70757a; text-align: left; }
        @media only screen and (max-width: 600px) {
          .wrapper { padding: 0; }
          .container { border-radius: 0; }
          .header, .content, .brand-footer { padding: 24px; }
          .detail-label { width: 140px; }
        }
      </style>
    </head>
    <body>
      <div class="wrapper">
        <div class="container">
          <div class="header">
            <div class="header-top">GOOGLE PLAY • ESCALATIONS</div>
            <h1 class="header-title">${config.title}</h1>
            <div class="header-icon">
              <img src="${iconUrl}" width="36" height="32" style="filter: brightness(0) invert(1);">
            </div>
          </div>
          <div class="content">
            <div class="message">${config.message}</div>
            <div class="details-box">
              <div class="details-header">CASE DETAILS</div>
              ${config.detailsHtml}
            </div>
            ${config.buttonUrl ? `
              <div class="button-container">
                <a href="${config.buttonUrl}" class="button">${config.buttonText || 'View Case'}</a>
              </div>
            ` : ''}
            ${config.footerNote ? `<div class="footer-note">${config.footerNote}</div>` : ''}
          </div>
          <div class="brand-footer">
            This is an automated message from the Google Play Escalations Dashboard. Please do not reply to this email.
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
}

function _detailRow(label, value, isLink) {
  const valHtml = isLink ? `<a href="${value}" class="detail-link" target="_blank">Open Case</a>` : value;
  return `
    <div class="detail-row">
      <div class="detail-label">${label}</div>
      <div class="detail-value">${valHtml}</div>
    </div>`;
}

function setupSLATrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkSLAWarnings') { ScriptApp.deleteTrigger(trigger); }
  });
  ScriptApp.newTrigger('checkSLAWarnings').timeBased().everyHours(1).create();
}
