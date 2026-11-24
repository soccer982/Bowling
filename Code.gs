/**
 * Alley Crew HQ — Google Apps Script Real-Time Sync Backend + 48h Reminders
 * Bound to: Alley Crew — Bowling Stats & Standings
 * Sheet ID: 1zlHenHgVtew6U5pxFjq_ZLcnFHQbXq6ecR36Y7HjTS8
 */

const SHEET_ID = '1zlHenHgVtew6U5pxFjq_ZLcnFHQbXq6ecR36Y7HjTS8';
const STATE_TAB = '_AppState';
const GAME_LOG_TAB = 'Game Log';

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  } catch (e) {
    return SpreadsheetApp.openById(SHEET_ID);
  }
}

function getOrCreateStateSheet_(ss) {
  let sh = ss.getSheetByName(STATE_TAB);
  if (!sh) {
    sh = ss.insertSheet(STATE_TAB);
    sh.hideSheet();
    sh.getRange('A1:B1').setValues([['Version', 'StateJSON']]);
    sh.getRange('A2:B2').setValues([[0, '']]);
  }
  return sh;
}

/**
 * Serves Index.html with optional ?player=<id> personalized login parameter
 * and allows embedding inside Google Sites.
 */
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.initialPlayer = (e && e.parameter && e.parameter.player) ? String(e.parameter.player) : '';
  return tpl.evaluate()
    .setTitle('Alley Crew HQ — Bowling Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getLiveState() {
  const ss = getSpreadsheet_();
  const sh = getOrCreateStateSheet_(ss);
  const row = sh.getRange('A2:B2').getValues()[0];
  return {
    version: Number(row[0] || 0),
    stateJson: String(row[1] || '')
  };
}

function saveLiveState(stateJson) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSpreadsheet_();
    const sh = getOrCreateStateSheet_(ss);
    const curVer = Number(sh.getRange('A2').getValue() || 0);
    const nextVer = curVer + 1;
    sh.getRange('A2:B2').setValues([[nextVer, stateJson]]);
    return { ok: true, version: nextVer };
  } finally {
    lock.releaseLock();
  }
}

function appendGameLogRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return { ok: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(GAME_LOG_TAB);
    if (!sh) return { ok: false, error: 'Game Log tab not found' };

    const colA = sh.getRange('A5:A500').getValues();
    let startRow = 5;
    for (let i = 0; i < colA.length; i++) {
      if (colA[i][0] === '' || colA[i][0] === null) { startRow = 5 + i; break; }
    }

    const formatted = rows.map((r, idx) => {
      const rowNum = startRow + idx;
      return [
        r.date, r.gameNum || 1, r.gameStyle, r.bowler, r.partner || '',
        Number(r.scratch || 0), Number(r.handicap || 0), `=F${rowNum}+G${rowNum}`,
        Number(r.strikes || 0), Number(r.spares || 0), Number(r.gutters || 0), Number(r.splits || 0),
        r.isWinner ? 'Y' : 'N', r.notes || ''
      ];
    });
    sh.getRange(startRow, 1, formatted.length, 14).setValues(formatted);
    return { ok: true, insertedAt: startRow, count: formatted.length };
  } finally {
    lock.releaseLock();
  }
}

function appendTeamMatchRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return { ok: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName('Team Handicaps & Matches');
    if (!sh) return { ok: false, error: 'Team Handicaps & Matches tab not found' };

    const colA = sh.getRange('A1:A500').getValues();
    let headerRow = -1;
    for (let i = 0; i < colA.length; i++) {
      if (String(colA[i][0]).trim() === 'Date') { headerRow = i + 1; break; }
    }
    if (headerRow === -1) return { ok: false, error: 'Team Match Log header not found' };

    let startRow = headerRow + 1;
    for (let i = headerRow; i < colA.length; i++) {
      if (colA[i][0] === '' || colA[i][0] === null) { startRow = i + 1; break; }
    }

    const formatted = rows.map((r, idx) => {
      const rowNum = startRow + idx;
      return [
        r.date, r.gameStyle, r.teamPair, r.opponent || '',
        Number(r.rawTeam || 0), Number(r.teamHcp || 0), `=E${rowNum}+F${rowNum}`,
        r.won ? 'W' : 'L', r.notes || ''
      ];
    });
    sh.getRange(startRow, 1, formatted.length, 9).setValues(formatted);
    return { ok: true, insertedAt: startRow, count: formatted.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Run this ONCE in Apps Script (Select installReminderTrigger -> Run)
 * to automatically check every 6 hours and email opted-in bowlers 48 hours before bowling!
 */
function installReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'send48HourReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('send48HourReminders').timeBased().everyHours(6).create();
}

/**
 * Checks the scheduled outing date & time; if within 48 hours, emails everyone
 * who linked an email address AND opted in to 48-hour reminders.
 */
function send48HourReminders() {
  const live = getLiveState();
  if (!live.stateJson) return;
  const state = JSON.parse(live.stateJson);
  if (!state.selectedDateId || !Array.isArray(state.members)) return;

  const dateKey = state.selectedDateId;
  if (state.remindersSentForDates && state.remindersSentForDates[dateKey]) return;

  const [y, m, d] = dateKey.split('-').map(Number);
  const [hh, mm] = (state.setTime || '19:00').split(':').map(Number);
  const eventTime = new Date(y, m - 1, d, hh, mm, 0);
  const hoursUntil = (eventTime.getTime() - Date.now()) / (1000 * 60 * 60);

  if (hoursUntil > 0 && hoursUntil <= 48) {
    const baseUrl = (state.webAppUrl || '').trim() || ScriptApp.getService().getUrl();
    const inNames = state.members
      .filter(p => p.rsvps && p.rsvps[dateKey] === 'in')
      .map(p => p.name);

    state.members.forEach(p => {
      if (p.email && p.remind48h) {
        const personalLink = baseUrl ? `${baseUrl}?player=${encodeURIComponent(p.id)}` : '';
        const subject = `🎳 Bowling Reminder: Alley Crew at Medford Lanes (${dateKey})`;
        const body = [
          `Hey ${p.name},`,
          ``,
          `Quick 48-hour heads-up — Alley Crew is scheduled to bowl at Medford Lanes (665 Stokes Rd, Medford, NJ) on ${dateKey} at ${state.setTime || '19:00'}.`,
          `Currently checked IN (${inNames.length}): ${inNames.join(', ') || 'None yet'}.`,
          ``,
          personalLink ? `Update your RSVP or game vote with your personal link:\n${personalLink}` : `Open the Alley Crew dashboard to confirm your RSVP and game vote.`,
          ``,
          `See you on the lanes!`
        ].join('\n');
        MailApp.sendEmail(p.email, subject, body);
      }
    });

    if (!state.remindersSentForDates) state.remindersSentForDates = {};
    state.remindersSentForDates[dateKey] = new Date().toISOString();
    saveLiveState(JSON.stringify(state));
  }
}
