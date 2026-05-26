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

/* ═══════════════════════════════════════════════════════════════════
   COLUMN MAP — Open Sup Cases (1-based)
   A=1  Timestamp
   B=2  Email (submitter)
   C=3  LDAP
   D=4  Case ID
   E=5  Symptom/Issue
   F=6  Detailed Customer Issue
   G=7  Reason for Escalation
   H=8  Channel
   I=9  Team
   J=10 Case ID Link
   K=11 Date          ← formula-populated, DO NOT write
   L=12 Time          ← formula-populated, DO NOT write
   M=13 Time spent before taken  (resolutionTime written on resolve)
   N=14 Handled by
   O=15 SME Remarks
   P=16 Claimed At    ← for TAT fix (B1), written by claimCase
   ═══════════════════════════════════════════════════════════════════

   COLUMN MAP — SupervisorList (1-based)
   A=1  Email Address
   B=2  LDAP
   C=3  Role  (Supervisor | SME)
   D=4+ Date columns  (header = "Apr-21", values = shift start time | "OFF" | "VL")
   ═══════════════════════════════════════════════════════════════════ */


/* ═══════════════════════════════════════════════════════════════════
   RBAC CORE
   ═══════════════════════════════════════════════════════════════════ */

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


/* ═══════════════════════════════════════════════════════════════════
   ROUTING & SESSION
   ═══════════════════════════════════════════════════════════════════ */

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


/* ═══════════════════════════════════════════════════════════════════
   SUBMIT
   ═══════════════════════════════════════════════════════════════════ */

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

    if (!/^\d-\d{12}$/.test(formData.caseId.trim())) return { success: false, message: 'Invalid Case ID format.' };

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const allData = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
      // COL_LDAP=2(C), COL_CASEID=3(D), COL_SYMPTOM=4(E) → 0-based: 2,3,4
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


/* ═══════════════════════════════════════════════════════════════════
   READ FUNCTIONS (available to all roles)
   ═══════════════════════════════════════════════════════════════════ */

/* ═══════════════════════════════════════════════════════════════════
   READ FUNCTIONS (Upgraded for Enterprise Speed & Caching)
   ═══════════════════════════════════════════════════════════════════ */

function getOpenCases() {
  _checkRateLimit('getOpenCases');
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName('Open Sup Cases');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const data = sheet.getRange(3, 1, lastRow - 2, 16).getValues();
  const openCases = [];
  const myEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  const myLdap = myEmail.split('@')[0];

  // BUILD INDEX: rowIdx → case object (skip resolved rows immediately)
  const INDEX_REMARKS = 14;  // col O (0-based)
  const INDEX_HANDLED = 13;  // col N (0-based)

  data.forEach((row, index) => {
    const remarks   = String(row[INDEX_REMARKS] || '').trim();
    if (remarks !== '') return; // indexed skip — resolved

    const handledBy = String(row[INDEX_HANDLED] || '').trim();
    const claimLdap = handledBy.toLowerCase().split('@')[0];
    const isMine    = (handledBy.toLowerCase() === myEmail) ||
                      (claimLdap === myLdap && claimLdap !== '');

    openCases.push({
      rowIdx        : index + 3,
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
  if (lastRow < 3) return [];

  // Memory Chunking: Only pull the last 3000 rows for metrics/history
  // This keeps the app lightning fast without breaking memory limits!
  const rowLimit = Math.min(lastRow - 2, 3000);
  const startRow = lastRow - rowLimit + 1;

  const data = sheet.getRange(startRow, 1, rowLimit, 17).getValues();
  const resolvedCases = [];

  // Loop backwards to get newest first
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


/* ═══════════════════════════════════════════════════════════════════
   SCHEDULE TRACKER
   Reads SupervisorList and returns today + next 2 days for each person.
   Date headers start at col D (col index 4, 1-based).
   Returns:
   [{ ldap, email, role, days: [{ date, label, value, status }] }]
   status: 'on'  = has a shift time
           'off' = "OFF"
           'vl'  = "VL"
           'unknown' = no column found for that date
   ═══════════════════════════════════════════════════════════════════ */

function getSchedule() {
  _checkRateLimit('getSchedule');
  try {
    const ss    = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet = ss.getSheetByName(SHEET_SUPERS);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 4) return [];
    Logger.log('lastCol detected: ' + lastCol); // temporary debug line

    // Read header row
    const headerRange = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
    // Read all data rows
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

    // Build the 3 target dates: today, tomorrow, day after
    const PHT_OFFSET_MS = 8 * 60 * 60 * 1000; // UTC+8
    const nowPHT        = new Date(Date.now() + PHT_OFFSET_MS);

    const targetDates = [0, 1, 2].map(function(offset) {
      const d = new Date(nowPHT);
      d.setUTCDate(d.getUTCDate() + offset);
      return d;
    });

    // Build a label matcher: "Apr-21" format
    function makeDateLabel(d) {
      const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return months[d.getUTCMonth()] + '-' + d.getUTCDate();
    }

    // Find column indices for each target date
      const targetColIndices = targetDates.map(function(d) {
      const lbl = makeDateLabel(d);
      
      const idx = headerRange.findIndex(function(h) {
        if (h instanceof Date) {
          // Convert sheet date to PHT to compare correctly
          const hPHT = new Date(h.getTime() + PHT_OFFSET_MS);
          const hMonth = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][hPHT.getUTCMonth()];
          const hDate = hPHT.getUTCDate();
          return (hMonth + '-' + hDate) === lbl;
        }
        return String(h).trim() === lbl;
      });
      
      return { date: d, label: lbl, colIdx: idx };
    });

    // Build display label like "Mon Apr 21"
    const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    function makeDisplayLabel(d, offset) {
      if (offset === 0) return 'Today';
      if (offset === 1) return 'Tomorrow';
      return DAY_NAMES[d.getUTCDay()] + ' ' + makeDateLabel(d).replace('-', ' ');
    }

    // Build result
    const result = [];
    dataRange.forEach(function(row) {
      const email = String(row[0] || '').trim();
      const ldap  = String(row[1] || '').trim();
      const role  = String(row[2] || '').trim();
      if (!ldap) return;

      const days = targetColIndices.map(function(tc, offset) {
        var rawVal = '';
        if (tc.colIdx >= 0) {
          // +3 offset: cols A,B,C are indices 0,1,2; col D starts at index 3
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
          // --- THE NEW 9-HOUR LIVE MATH ---
          status = 'scheduled'; // Default to a gray badge for inactive shifts
          displayValue = rawVal;

          if (offset === 0) { // Only do real-time math for "Today"
            
            // Helper to turn string "1:00:00 AM" into a rigid UTC timestamp
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

            // 1. Check Today's shift window
            var shiftToday = getShift(tc.date, rawVal);
            
            // 2. Check Yesterday's shift window (catches night shifts crossing midnight!)
            var yestVal = tc.colIdx > 0 ? String(row[tc.colIdx + 3 - 1] || '').trim() : '';
            var yestDate = new Date(tc.date.getTime() - 86400000); // Minus 24 hours
            var shiftYest = getShift(yestDate, yestVal);

            var isActive = false;
            if (shiftToday && nowPHT >= shiftToday.start && nowPHT < shiftToday.end) isActive = true;
            if (shiftYest && nowPHT >= shiftYest.start && nowPHT < shiftYest.end) isActive = true;

            if (isActive) status = 'on'; // If they are currently active, turn it Green!

          } else {
            status = 'on'; // Tomorrow and the Day After default to green
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


/* ═══════════════════════════════════════════════════════════════════
   METRICS SNAPSHOT
   Called hourly by a time-driven trigger.
   Writes one row to the "Metrics" sheet with current queue stats.
   ═══════════════════════════════════════════════════════════════════ */

/**
 * Writes a snapshot row to the Metrics sheet.
 * Safe to run manually from the Apps Script editor to test.
 */
function recordMetricsSnapshot() {
  try {
    const ss          = SpreadsheetApp.openById(MASTER_DB_ID);
    const openSheet   = ss.getSheetByName('Open Sup Cases');
    if (!openSheet) return;

    // ── Get or create Metrics sheet ──
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
      metricsSheet.setColumnWidth(1, 180);
      metricsSheet.setColumnWidth(7, 280);
    }

    // ── Read all Open Sup Cases rows ──
    const lastRow = openSheet.getLastRow();
    if (lastRow < 3) {
      metricsSheet.appendRow([new Date(), 0, 0, 0, 0, '', '—', '—', '—']);
      return;
    }

    const data = openSheet.getRange(3, 1, lastRow - 2, 17).getValues();

    // PHT midnight for "today"
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
      const handledBy    = String(row[13] || '').trim(); // col N
      const remarks      = String(row[14] || '').trim(); // col O
      const claimedAt    = row[15];                      // col P
      const resolutionTs = row[12];                      // col M

      if (remarks === '') {
        // Open case
        openCount++;
        if (handledBy !== '') claimedCount++;
      } else {
        // Resolved case — only count if resolved today (PHT)
        if (resolutionTs) {
          const resDate = new Date(resolutionTs);
          if (!isNaN(resDate.getTime()) && resDate.getTime() >= midnightPHT.getTime()) {
            resolvedToday++;

            // TAT: claimedAt → resolutionTs
            const startTs = claimedAt ? new Date(claimedAt) : new Date(row[0]);
            if (!isNaN(startTs.getTime())) {
              const diff = (resDate.getTime() - startTs.getTime()) / 60000;
              if (diff > 0) { tatTotalMins += diff; tatCount++; }
            }

            // Symptom tally
            const sym = String(row[4] || '').trim();
            if (sym) symptoms[sym] = (symptoms[sym] || 0) + 1;

            // Channel tally
            const chan = String(row[7] || '').trim();
            if (chan) channels[chan] = (channels[chan] || 0) + 1;

            // SME tally
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


/* ═══════════════════════════════════════════════════════════════════
   TRIGGER SETUP
   Run setupHourlyTrigger() ONCE from the Apps Script editor.
   It installs a time-driven trigger that calls recordMetricsSnapshot
   every hour. Running it again is safe — it deletes the old trigger
   first to avoid duplicates.
   ═══════════════════════════════════════════════════════════════════ */

/**
 * Run this ONCE from the Apps Script editor to install the hourly trigger.
 * Safe to re-run — removes any existing trigger for recordMetricsSnapshot first.
 */
function setupHourlyTrigger() {
  // Remove existing triggers for this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'recordMetricsSnapshot') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Install a new every-hour trigger
  ScriptApp.newTrigger('recordMetricsSnapshot')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ Hourly trigger for recordMetricsSnapshot installed successfully.');
}


/* ═══════════════════════════════════════════════════════════════════
   WRITE FUNCTIONS (SME / Supervisor only)
   ═══════════════════════════════════════════════════════════════════ */

function claimCase(rowIdx) {
  _checkRateLimit('claimCase');
  requireSME_();
  try {
    const ss        = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet     = ss.getSheetByName('Open Sup Cases');
    const smeEmail  = Session.getActiveUser().getEmail();
    const claimedAt = new Date();

    sheet.getRange(rowIdx, 16).setValue(claimedAt); // col 16 (P) = Claimed At
    sheet.getRange(rowIdx, 14).setValue(smeEmail);  // <-- PUT THIS BACK HERE
    _log('CLAIM', smeEmail, 'Case claimed', rowIdx, caseData.caseId, caseData.ldap);

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
    sheet.getRange(rowIdx, 15).setValue(remarks);      // ← ADD THIS between 650 and 651
    sheet.getRange(rowIdx, 17).setValue(resolutionType || '');
    _log('RESOLVE', smeEmail, 'Case resolved | Driver: ' + resolutionType, rowIdx, caseData.caseId, remarks.substring(0, 80));

    sendResolveNotification_(caseData, smeEmail);
    hideRows();

    return { success: true, message: 'Case resolved successfully!' };
  } catch (e) { return { success: false, message: e.message }; }
}


/* ═══════════════════════════════════════════════════════════════════
   EMAIL NOTIFICATIONS
   ═══════════════════════════════════════════════════════════════════ */

function sendClaimNotification_(caseData, smeEmail) {
  try {
    if (!caseData.submitter) return;

    const smeLdap    = smeEmail.split('@')[0];
    const subject    = '✅ Your escalation has been picked up — Case ' + caseData.caseId;
    const claimedAt  = new Date().toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' });
    const submittedAt = caseData.timestamp
      ? new Date(caseData.timestamp).toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' })
      : '—';

    const caseLinkHtml = caseData.caseLink
      ? '<a href="' + caseData.caseLink + '" style="color:#1A73E8;">Open Case</a>'
      : '—';

    const body = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:'Google Sans',Roboto,Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);max-width:600px;">
        <tr>
          <td style="background:linear-gradient(135deg,#1A73E8 0%,#0d47a1 100%);padding:28px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td><div style="color:#ffffff;font-size:11px;font-weight:600;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;">Google Play · Escalations</div>
              <div style="color:#ffffff;font-size:22px;font-weight:700;line-height:1.3;">Case Picked Up</div></td>
              <td align="right"><div style="background:rgba(255,255,255,0.2);border-radius:8px;padding:8px 14px;display:inline-block;"><span style="color:#ffffff;font-size:20px;">✅</span></div></td>
            </tr></table>
          </td>
        </tr>
        <tr>
          <td style="padding:28px 32px;">
            <p style="margin:0 0 20px;color:#3c4043;font-size:15px;line-height:1.6;">
              Hi <strong>${caseData.ldap}</strong>, your escalation has been picked up by <strong>${smeLdap}</strong> and is now <span style="color:#1A73E8;font-weight:600;">In Progress</span>.
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8f9fa;border-radius:8px;border-left:4px solid #1A73E8;margin-bottom:20px;">
              <tr><td style="padding:20px 24px;">
                <div style="color:#5f6368;font-size:11px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;">Case Details</div>
                ${row_('Case ID',      caseData.caseId)}
                ${row_('Symptom',      caseData.symptom)}
                ${row_('Channel',      caseData.channel)}
                ${row_('Team',         caseData.team)}
                ${row_('Case Link',    caseLinkHtml, true)}
                ${row_('Submitted',    submittedAt)}
                ${row_('Picked up at', claimedAt)}
                ${row_('Handled by',   smeLdap)}
              </td></tr>
            </table>
            <p style="margin:0;color:#5f6368;font-size:13px;line-height:1.6;">You will receive another notification once your case has been resolved. If you have urgent updates, please contact your supervisor directly.</p>
          </td>
        </tr>
        <tr><td style="background:#f8f9fa;padding:16px 32px;border-top:1px solid #e8eaed;"><p style="margin:0;color:#9aa0a6;font-size:12px;">This is an automated message from the Google Play Escalations Dashboard. Please do not reply to this email.</p></td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

    MailApp.sendEmail({
      to      : caseData.submitter,
      subject : subject,
      htmlBody: body,
      name    : 'Play Escalations · Google Play'
    });

  } catch (e) {
    console.error('sendClaimNotification_ error: ' + e.message);
  }
}

function sendResolveNotification_(caseData, smeEmail) {
  try {
    if (!caseData.submitter) return;

    const smeLdap     = smeEmail.split('@')[0];
    const subject     = '🎉 Your escalation has been resolved — Case ' + caseData.caseId;
    const resolvedAt  = new Date(caseData.resolutionTime).toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' });
    const submittedAt = caseData.timestamp
      ? new Date(caseData.timestamp).toLocaleString('en-US', { timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short' })
      : '—';

    const caseLinkHtml = caseData.caseLink
      ? '<a href="' + caseData.caseLink + '" style="color:#1A73E8;">Open Case</a>'
      : '—';

    const body = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:'Google Sans',Roboto,Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);max-width:600px;">
        <tr>
          <td style="background:linear-gradient(135deg,#34a853 0%,#1e7e34 100%);padding:28px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td><div style="color:#ffffff;font-size:11px;font-weight:600;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;">Google Play · Escalations</div>
              <div style="color:#ffffff;font-size:22px;font-weight:700;line-height:1.3;">Case Resolved</div></td>
              <td align="right"><div style="background:rgba(255,255,255,0.2);border-radius:8px;padding:8px 14px;display:inline-block;"><span style="color:#ffffff;font-size:20px;">🎉</span></div></td>
            </tr></table>
          </td>
        </tr>
        <tr>
          <td style="padding:28px 32px;">
            <p style="margin:0 0 20px;color:#3c4043;font-size:15px;line-height:1.6;">
              Hi <strong>${caseData.ldap}</strong>, your escalation has been <span style="color:#34a853;font-weight:600;">resolved</span> by <strong>${smeLdap}</strong>.
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8f9fa;border-radius:8px;border-left:4px solid #34a853;margin-bottom:20px;">
              <tr><td style="padding:20px 24px;">
                <div style="color:#5f6368;font-size:11px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;">Case Details</div>
                ${row_('Case ID',     caseData.caseId)}
                ${row_('Symptom',     caseData.symptom)}
                ${row_('Channel',     caseData.channel)}
                ${row_('Team',        caseData.team)}
                ${row_('Case Link',   caseLinkHtml, true)}
                ${row_('Submitted',   submittedAt)}
                ${row_('Resolved at', resolvedAt)}
                ${row_('Resolved by', smeLdap)}
              </td></tr>
            </table>
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#e8f5e9;border-radius:8px;border-left:4px solid #34a853;margin-bottom:20px;">
              <tr><td style="padding:20px 24px;">
                <div style="color:#2e7d32;font-size:11px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:10px;">SME Remarks</div>
                <div style="color:#3c4043;font-size:14px;line-height:1.7;white-space:pre-wrap;">${escHtml_(caseData.remarks)}</div>
              </td></tr>
            </table>
            <p style="margin:0;color:#5f6368;font-size:13px;line-height:1.6;">If you have any follow-up concerns, please contact your supervisor or submit a new escalation request.</p>
          </td>
        </tr>
        <tr><td style="background:#f8f9fa;padding:16px 32px;border-top:1px solid #e8eaed;"><p style="margin:0;color:#9aa0a6;font-size:12px;">This is an automated message from the Google Play Escalations Dashboard. Please do not reply to this email.</p></td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

    MailApp.sendEmail({
      to      : caseData.submitter,
      subject : subject,
      htmlBody: body
    });

  } catch (e) {
    console.error('sendResolveNotification_ error: ' + e.message);
  }
}

function row_(label, value, isHtml) {
  const displayValue = isHtml ? (value || '—') : escHtml_(value || '—');
  return `
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px;">
      <tr>
        <td width="38%" style="color:#5f6368;font-size:13px;padding-right:8px;vertical-align:top;">${label}</td>
        <td style="color:#3c4043;font-size:13px;font-weight:500;vertical-align:top;">${displayValue}</td>
      </tr>
    </table>`;
}

function escHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


/* ═══════════════════════════════════════════════════════════════════
   UTILITIES
   ═══════════════════════════════════════════════════════════════════ */

// ── AUDIT LOG HELPERS ─────────────────────────────────────────────
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
      logSheet.appendRow([
        'Timestamp','Actor Email','Actor LDAP','Action',
        'Target Row','Case ID','Detail','Extra'
      ]);
      logSheet.getRange(1, 1, 1, 8)
        .setFontWeight('bold')
        .setBackground('#1A73E8')
        .setFontColor('#FFFFFF');
      logSheet.setFrozenRows(1);
      logSheet.setColumnWidth(1, 180);
      logSheet.setColumnWidth(7, 320);
    }

    const actorLdap = String(actor || '').split('@')[0];
    logSheet.appendRow([
      new Date(),
      actor      || '',
      actorLdap  || '',
      action     || '',
      targetRow  != null ? targetRow : '',
      caseId     || '',
      detail     || '',
      extra      || ''
    ]);
  } catch(e) {
    console.error('_log failed: ' + e.message);
  }
}

function hideRows() {
  const ss    = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName("Open Sup Cases");
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const values = sheet.getRange(3, 15, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      sheet.hideRows(3 + i);
    }
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
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
function testEmail() {
  MailApp.sendEmail({
    to: 'stevenjosephc@google.com',
    subject: 'Test',
    body: 'Email works!'
  });
}
/* ═══════════════════════════════════════════════════════════════════
   SLA WARNING EMAILS
   Runs hourly via trigger. Sends a 24hr reminder and a 48hr urgent
   warning for any open case that has not yet been resolved.
═══════════════════════════════════════════════════════════════════ */

var TEAM_SUPERVISOR_MAP = {
  'Team Steven': 'stevenjosephc',
  'Team Gerry':  'gerrymae',
  'Team Khent':  'khent',
  'Team James':  'jamessevilla',
  'Team Denden': 'bernardboy',
  'Team Jim':    'jadoptante',
  'Team Al':     'acaluang',
  'Team Faye':   'fajirnah',
  'Team Mel':    'tapalla',
  'Team Mary':   'marineth'
};

var TEAM_SME_MAP = {
  'Team Steven': ['sheenamae'],
  'Team Gerry':  ['caval'],
  'Team Khent':  ['criseldaa'],
  'Team James':  ['acemile'],
  'Team Denden': ['jgarlet'],
  'Team Jim':    ['glendajoe'],
  'Team Al':     ['elagarto'],
  'Team Faye':   ['neljhon', 'mquirol'],
  'Team Mel':    ['fykeivan', 'mabubay'],
  'Team Mary':   ['fykeivan']
};

var SLA_CC_48HR = ['deanmark@google.com', 'joliveros@google.com', 'jmontrias@google.com'];

function checkSLAWarnings() {
  try {
    const ss        = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet     = ss.getSheetByName('Open Sup Cases');
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    // Read cols A–S (1–19)
    const data    = sheet.getRange(3, 1, lastRow - 2, 19).getValues();
    const now     = new Date();
    const MS_24HR = 24 * 60 * 60 * 1000;
    const MS_48HR = 48 * 60 * 60 * 1000;

    // Col indices (0-based)
    const COL_TIMESTAMP   = 0;   // A
    const COL_LDAP        = 2;   // C
    const COL_CASEID      = 3;   // D
    const COL_SYMPTOM     = 4;   // E
    const COL_REASON      = 6;   // G
    const COL_CHANNEL     = 7;   // H
    const COL_TEAM        = 8;   // I
    const COL_REMARKS     = 14;  // O
    const COL_WARN24      = 17;  // R
    const COL_WARN48      = 18;  // S

    data.forEach(function(row, i) {
      const remarks   = String(row[COL_REMARKS] || '').trim();
      if (remarks !== '') return; // already resolved, skip

      const timestamp = row[COL_TIMESTAMP];
      if (!timestamp) return;

      const submitted = new Date(timestamp);
      if (isNaN(submitted.getTime())) return;

      const ageMs       = now.getTime() - submitted.getTime();
      const warn24Sent  = row[COL_WARN24] === true || String(row[COL_WARN24]).toUpperCase() === 'TRUE';
      const warn48Sent  = row[COL_WARN48] === true || String(row[COL_WARN48]).toUpperCase() === 'TRUE';

      const ldap        = String(row[COL_LDAP]    || '').trim();
      const caseId      = String(row[COL_CASEID]  || '').trim();
      const symptom     = String(row[COL_SYMPTOM] || '').trim();
      const reason      = String(row[COL_REASON]  || '').trim();
      const channel     = String(row[COL_CHANNEL] || '').trim();
      const team        = String(row[COL_TEAM]    || '').trim();
      const sheetRow    = i + 3; // actual sheet row number

      const agentEmail  = ldap + '@google.com';
      const supLdap     = TEAM_SUPERVISOR_MAP[team] || '';
      const supEmail    = supLdap ? supLdap + '@google.com' : '';
      const smeLdaps    = TEAM_SME_MAP[team] || [];
      const smeEmails   = smeLdaps.map(function(s) { return s + '@google.com'; });
      const submittedStr = submitted.toLocaleString('en-US', {
        timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short'
      });
      const hoursOld    = Math.floor(ageMs / (1000 * 60 * 60));

      // ── 48HR WARNING (check first so it doesn't double-send 24hr) ──
      if (ageMs >= MS_48HR && !warn48Sent) {
        sendSLA48hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, symptom, reason, channel, team, submittedStr, hoursOld);
        sheet.getRange(sheetRow, COL_WARN24 + 1).setValue(true); // mark 24 too
        sheet.getRange(sheetRow, COL_WARN48 + 1).setValue(true);
        Logger.log('[SLA 48hr] ' + caseId + ' | ' + ldap);
        _log('SLA_48HR', 'system', 'SLA 48hr warning sent', sheetRow, caseId, ldap); // <-- INSERTED HERE (Line 1147)
        return;
      }

      // ── 24HR WARNING ──
      if (ageMs >= MS_24HR && !warn24Sent) {
        sendSLA24hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, symptom, reason, channel, team, submittedStr, hoursOld);
        sheet.getRange(sheetRow, COL_WARN24 + 1).setValue(true);
        _log('SLA_24HR', 'system', 'SLA 24hr warning sent', sheetRow, caseId, ldap); // <-- INSERTED HERE (Line 1155)
      }
    });

  } catch(e) {
    console.error('checkSLAWarnings error: ' + e.message);
  }
}

/* ─────────────────────────────────────────
   24HR EMAIL — Friendly Reminder
───────────────────────────────────────── */
function sendSLA24hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, symptom, reason, channel, team, submittedStr, hoursOld, testMode_) {
  try {
    const subject = '⏰ Friendly Reminder: Your Escalation Has Been Pending for 24 Hours';
    const caseLink = 'https://cases.connect.corp.google.com/' + caseId;

    const body = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:'Google Sans',Roboto,Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);max-width:600px;">

        <!-- HEADER -->
        <tr>
          <td style="background:linear-gradient(135deg,#F9AB00 0%,#E37400 100%);padding:28px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <div style="display:flex;align-items:center;gap:12px;">
                  <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_Play_2022_icon.svg/960px-Google_Play_2022_icon.svg.png"
                       width="40" height="40" style="border-radius:8px;display:block;" alt="Google Play">
                </div>
                <div style="color:#ffffff;font-size:11px;font-weight:600;letter-spacing:2px;text-transform:uppercase;margin-top:12px;opacity:0.85;">Google Play · Escalations</div>
                <div style="color:#ffffff;font-size:22px;font-weight:700;line-height:1.3;margin-top:4px;">24-Hour Pending Reminder</div>
                <div style="color:rgba(255,255,255,0.8);font-size:13px;margin-top:6px;">Your escalation is still awaiting supervisor action.</div>
              </td>
              <td align="right" valign="top">
                <div style="background:rgba(255,255,255,0.2);border-radius:12px;padding:14px 18px;text-align:center;display:inline-block;">
                  <div style="font-size:36px;line-height:1;">⏰</div>
                  <div style="color:#fff;font-size:11px;font-weight:700;margin-top:4px;letter-spacing:.5px;">${hoursOld}h OLD</div>
                </div>
              </td>
            </tr></table>
          </td>
        </tr>

        <!-- STRIPE -->
        <tr><td style="height:4px;background:linear-gradient(90deg,#4285F4,#EA4335,#FBBC05,#34A853)"></td></tr>

        <!-- BODY -->
        <tr>
          <td style="padding:28px 32px;">
            <p style="margin:0 0 20px;color:#3c4043;font-size:15px;line-height:1.7;">
              Hi <strong>${ldap}</strong>, this is a friendly reminder that your escalation submitted <strong>${submittedStr}</strong> has been in the queue for over <strong>${hoursOld} hours</strong> and is still awaiting action from a supervisor or SME.
            </p>

            <!-- CASE DETAILS BOX -->
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#FEF7E0;border-radius:10px;border-left:4px solid #F9AB00;margin-bottom:24px;">
              <tr><td style="padding:20px 24px;">
                <div style="color:#B06000;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;">📋 Case Details</div>
                ${slaRow_('Case ID',   '<a href="' + caseLink + '" style="color:#1A73E8;font-weight:600;">' + caseId + '</a>', true)}
                ${slaRow_('Agent',     ldap)}
                ${slaRow_('Symptom',   symptom)}
                ${slaRow_('Channel',   channel)}
                ${slaRow_('Team',      team)}
                ${slaRow_('Submitted', submittedStr)}
                ${slaRow_('Age',       hoursOld + ' hours in queue')}
              </td></tr>
            </table>

            <!-- WHAT TO DO -->
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#E8F0FE;border-radius:10px;margin-bottom:24px;">
              <tr><td style="padding:18px 24px;">
                <div style="color:#1967D2;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;">💡 What happens next?</div>
                <p style="margin:0;color:#3c4043;font-size:13px;line-height:1.7;">
                  A supervisor or SME from <strong>${team}</strong> will pick up your case shortly. If this is urgent, please reach out to your team lead directly. You will receive another notification if your case remains unresolved at the 48-hour mark.
                </p>
              </td></tr>
            </table>

            <p style="margin:0;color:#5f6368;font-size:12px;line-height:1.6;border-top:1px solid #e8eaed;padding-top:16px;">
              This is an automated reminder from the Google Play Escalations system. Please do not reply to this email. If you believe this case has already been resolved, no further action is needed.
            </p>
          </td>
        </tr>

        <!-- FOOTER -->
        <tr>
          <td style="background:#f8f9fa;padding:16px 32px;border-top:1px solid #e8eaed;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_Play_2022_icon.svg/960px-Google_Play_2022_icon.svg.png"
                     width="20" height="20" style="display:inline-block;vertical-align:middle;margin-right:8px;" alt="">
                <span style="color:#9aa0a6;font-size:12px;vertical-align:middle;">Google Play Escalations · Automated Notification · Do not reply</span>
              </td>
            </tr></table>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;

    var ccList = [];
    if (!testMode_) {
      if (supEmail) ccList.push(supEmail);
      smeEmails.forEach(function(e) { if (e) ccList.push(e); });
    }

    MailApp.sendEmail({
      to:       agentEmail,
      cc:       ccList.filter(Boolean).join(','),
      subject:  subject,
      htmlBody: body,
      name:     'Play Escalations · Google Play',
      noReply:  true
    });

  } catch(e) {
    console.error('sendSLA24hrEmail_ error: ' + e.message);
  }
}

/* ─────────────────────────────────────────
   48HR EMAIL — Urgent Final Warning
───────────────────────────────────────── */
function sendSLA48hrEmail_(agentEmail, supEmail, smeEmails, caseId, ldap, symptom, reason, channel, team, submittedStr, hoursOld, testMode_) {
  try {
    const subject = '🚨 Urgent: Escalation Case Unresolved for 48 Hours — Immediate Action Required';
    const caseLink = 'https://cases.connect.corp.google.com/' + caseId;

    const body = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:'Google Sans',Roboto,Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);max-width:600px;">

        <!-- HEADER -->
        <tr>
          <td style="background:linear-gradient(135deg,#C5221F 0%,#EA4335 100%);padding:28px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_Play_2022_icon.svg/960px-Google_Play_2022_icon.svg.png"
                     width="40" height="40" style="border-radius:8px;display:block;" alt="Google Play">
                <div style="color:#ffffff;font-size:11px;font-weight:600;letter-spacing:2px;text-transform:uppercase;margin-top:12px;opacity:0.85;">Google Play · Escalations · Urgent Alert</div>
                <div style="color:#ffffff;font-size:22px;font-weight:700;line-height:1.3;margin-top:4px;">48-Hour SLA Breach Warning</div>
                <div style="color:rgba(255,255,255,0.85);font-size:13px;margin-top:6px;">This case has exceeded the 48-hour response threshold and requires immediate attention.</div>
              </td>
              <td align="right" valign="top">
                <div style="background:rgba(255,255,255,0.15);border:2px solid rgba(255,255,255,0.4);border-radius:12px;padding:14px 18px;text-align:center;display:inline-block;">
                  <div style="font-size:36px;line-height:1;">🚨</div>
                  <div style="color:#fff;font-size:11px;font-weight:700;margin-top:4px;letter-spacing:.5px;">${hoursOld}h OLD</div>
                </div>
              </td>
            </tr></table>
          </td>
        </tr>

        <!-- STRIPE -->
        <tr><td style="height:4px;background:linear-gradient(90deg,#EA4335,#EA4335,#FBBC05,#EA4335)"></td></tr>

        <!-- URGENT BANNER -->
        <tr>
          <td style="background:#FCE8E6;padding:16px 32px;border-bottom:1px solid #F5C6C6;">
            <p style="margin:0;color:#C5221F;font-size:14px;font-weight:600;text-align:center;">
              ⚠️ &nbsp;This case has been in the queue for <strong>${hoursOld} hours</strong> without resolution. Management has been notified.
            </p>
          </td>
        </tr>

        <!-- BODY -->
        <tr>
          <td style="padding:28px 32px;">
            <p style="margin:0 0 20px;color:#3c4043;font-size:15px;line-height:1.7;">
              Hi <strong>${ldap}</strong>, your escalation submitted on <strong>${submittedStr}</strong> has now been pending for over <strong>${hoursOld} hours</strong>. This exceeds the acceptable response threshold. Your team managers have been notified and this case requires <strong style="color:#C5221F;">immediate action</strong>.
            </p>

            <!-- CASE DETAILS BOX -->
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#FCE8E6;border-radius:10px;border-left:4px solid #EA4335;margin-bottom:24px;">
              <tr><td style="padding:20px 24px;">
                <div style="color:#C5221F;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;">🚨 Case Details</div>
                ${slaRow_('Case ID',   '<a href="' + caseLink + '" style="color:#C5221F;font-weight:700;">' + caseId + '</a>', true)}
                ${slaRow_('Agent',     ldap)}
                ${slaRow_('Symptom',   symptom)}
                ${slaRow_('Reason',    reason)}
                ${slaRow_('Channel',   channel)}
                ${slaRow_('Team',      team)}
                ${slaRow_('Submitted', submittedStr)}
                ${slaRow_('Age',       '<strong style="color:#C5221F;">' + hoursOld + ' hours — SLA BREACHED</strong>', true)}
              </td></tr>
            </table>

            <!-- REQUIRED ACTION -->
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#FEF7E0;border-radius:10px;border-left:4px solid #F9AB00;margin-bottom:24px;">
              <tr><td style="padding:18px 24px;">
                <div style="color:#B06000;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;">⚡ Required Action</div>
                <p style="margin:0 0 10px;color:#3c4043;font-size:13px;line-height:1.7;">
                  The following actions must be taken immediately:
                </p>
                <ul style="margin:0;padding-left:20px;color:#3c4043;font-size:13px;line-height:2;">
                  <li>An SME or Supervisor from <strong>${team}</strong> must claim and resolve this case immediately.</li>
                  <li>If the assigned team is unavailable, please escalate to your manager directly.</li>
                  <li>Update the case in the <a href="https://cases.connect.corp.google.com/${caseId}" style="color:#1A73E8;">Escalations Dashboard</a> once resolved.</li>
                </ul>
              </td></tr>
            </table>

            <!-- MANAGERS NOTIFIED -->
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#E6F4EA;border-radius:10px;border-left:4px solid #34A853;margin-bottom:24px;">
              <tr><td style="padding:16px 24px;">
                <div style="color:#137333;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">✅ Managers Notified</div>
                <p style="margin:0;color:#3c4043;font-size:13px;line-height:1.6;">
                  This alert has been automatically copied to the Play Ops management team for visibility and follow-up.
                </p>
              </td></tr>
            </table>

            <p style="margin:0;color:#5f6368;font-size:12px;line-height:1.6;border-top:1px solid #e8eaed;padding-top:16px;">
              This is an automated urgent alert from the Google Play Escalations system. Please do not reply to this email. This notification was generated because the case exceeded the 48-hour SLA threshold.
            </p>
          </td>
        </tr>

        <!-- FOOTER -->
        <tr>
          <td style="background:#f8f9fa;padding:16px 32px;border-top:1px solid #e8eaed;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_Play_2022_icon.svg/960px-Google_Play_2022_icon.svg.png"
                     width="20" height="20" style="display:inline-block;vertical-align:middle;margin-right:8px;" alt="">
                <span style="color:#9aa0a6;font-size:12px;vertical-align:middle;">Google Play Escalations · Automated Urgent Alert · Do not reply</span>
              </td>
              <td align="right">
                <span style="color:#EA4335;font-size:11px;font-weight:700;">SLA BREACH · ${hoursOld}h</span>
              </td>
            </tr></table>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;

    var ccList = [];
    if (supEmail) ccList.push(supEmail);
    smeEmails.forEach(function(e) { if (e) ccList.push(e); });
    if (!testMode_) {
      SLA_CC_48HR.forEach(function(e) { ccList.push(e); });
    }

    MailApp.sendEmail({
      to:       agentEmail,
      cc:       ccList.filter(Boolean).join(','),
      subject:  subject,
      htmlBody: body,
      name:     'Play Escalations · Google Play',
      noReply:  true
    });

  } catch(e) {
    console.error('sendSLA48hrEmail_ error: ' + e.message);
  }
}

/* ─────────────────────────────────────────
   EMAIL ROW HELPER
───────────────────────────────────────── */
function slaRow_(label, value, isHtml) {
  const displayValue = isHtml ? (value || '—') : escHtml_(value || '—');
  return `
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px;">
      <tr>
        <td width="32%" style="color:#5f6368;font-size:12px;padding-right:8px;vertical-align:top;font-weight:500;">${label}</td>
        <td style="color:#3c4043;font-size:13px;font-weight:500;vertical-align:top;">${displayValue}</td>
      </tr>
    </table>`;
}

/* ─────────────────────────────────────────
   TRIGGER SETUP — Run once from editor
───────────────────────────────────────── */
function setupSLATrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkSLAWarnings') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkSLAWarnings')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ SLA Warning trigger installed successfully.');
}
function testSLAEmails() {
  var ME          = 'stevenjosephc@google.com';
  var caseId      = '4-9048000040554';
  var ldap        = 'testldap';
  var symptom     = 'Can\'t redeem gift card';
  var reason      = 'Customer is requesting supervisor intervention.';
  var channel     = 'Chat';
  var team        = 'Team Steven';
  var submittedStr = new Date().toLocaleString('en-US', {
    timeZone: 'Asia/Manila', dateStyle: 'medium', timeStyle: 'short'
  });

  // Route ALL recipients to just you — no one else gets anything
  sendSLA24hrEmail_(ME, '', [], caseId, ldap, symptom, reason, channel, team, submittedStr, 24, true);
  sendSLA48hrEmail_(ME, '', [], caseId, ldap, symptom, reason, channel, team, submittedStr, 48, true);

  Logger.log('✅ Test emails sent only to ' + ME);
}
function initSLAColumns() {
  const ss    = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName('Open Sup Cases');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const data = sheet.getRange(3, 1, lastRow - 2, 19).getValues();
  const now  = new Date();
  const MS_24HR = 24 * 60 * 60 * 1000;
  const MS_48HR = 48 * 60 * 60 * 1000;

  let marked24 = 0, marked48 = 0;

  data.forEach(function(row, i) {
    const remarks  = String(row[14] || '').trim(); // col O
    if (remarks !== '') return; // already resolved, skip

    const timestamp = row[0]; // col A
    if (!timestamp) return;

    const submitted = new Date(timestamp);
    if (isNaN(submitted.getTime())) return;

    const ageMs   = now.getTime() - submitted.getTime();
    const sheetRow = i + 3;

    if (ageMs >= MS_48HR) {
      sheet.getRange(sheetRow, 18).setValue(true); // R
      sheet.getRange(sheetRow, 19).setValue(true); // S
      marked48++;
    } else if (ageMs >= MS_24HR) {
      sheet.getRange(sheetRow, 18).setValue(true); // R
      marked24++;
    }
  });

  Logger.log('✅ initSLAColumns complete. Marked ' + marked24 + ' cases as 24hr warned, ' + marked48 + ' cases as 48hr warned.');
}

function debugSchedule() {
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheetByName('SupervisorList');
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  // Read header row as values (Date objects)
  const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  // Read header row as display text
  const headerDisp = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  Logger.log('lastRow: ' + lastRow + ' | lastCol: ' + lastCol);

  for (var i = 3; i < headerVals.length; i++) {
    Logger.log('Col ' + (i+1) + ' | value type: ' + typeof headerVals[i] + 
               ' | value: ' + headerVals[i] + 
               ' | display: ' + headerDisp[i] +
               ' | isDate: ' + (headerVals[i] instanceof Date));
  }

  // Check today's date label
  const PHT_OFFSET_MS = 8 * 60 * 60 * 1000;
  const nowPHT = new Date(Date.now() + PHT_OFFSET_MS);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const todayLabel = months[nowPHT.getUTCMonth()] + '-' + nowPHT.getUTCDate();
  Logger.log('Today label we are looking for: ' + todayLabel);
}
