// ====================== SETTINGS ======================
const MASTER_URL = 'PASTE_MASTER_SHEET_URL_HERE';
const MASTER_TAB = 'PASTE_EXACT_TAB_NAME_HERE';

const MY_COURSES = ['PASTE_YOUR_COURSES'];
const SKIP_BUDGET_PCT = 0.20;

const EXAM_COURSE_MAP = {
  'Q1-COURSE': { course: 'COURSE', type: 'quiz' },
  'Q2-COURSE': { course: 'COURSE', type: 'quiz' },
  'ET-COURSE':  { course: 'COURSE',  type: 'exam' },
};

// Display columns for My Courses grid — -1 means lunch break
const DISPLAY_SLOTS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
// ========================================================

const DATE_COL_INDEX = 0;
const ROOM_COL_INDEX = 1;
const SLOT_COL_INDEXES = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
const SLOT_LABELS = [
  'S1  9:00-10:15am', 'S2  10:30-11:45am', 'S3  12:00-1:15pm',
  'Lunch/Break',
  'S4  2:30-3:45pm', 'S5  4:00-5:15pm', 'S6  5:30-6:45pm',
  'S7  7:00-8:15pm', 'S8  8:45-10:00pm', 'S9  10:15-11:30pm'
];
const LAST_COL_TO_READ = 12;

const SHEET_TODAY = "📌 Today's Classes";
const SHEET_ATTENDANCE = '✏️ Mark Attendance';
const SHEET_MY_COURSES = 'My Courses';
const SHEET_CONFLICTS = '⚠️ Conflicts';
const SHEET_WEEKLY = '📊 Weekly Summary';
const SHEET_DAYS_OFF = 'Personal Days Off';
const SHEET_ISSUES = '🩺 Issues Log';

const ATT_STATUS = { NOT_YET: 'Not yet', ATTENDED: 'Attended', SKIPPED: 'Skipped' };

const FONT = 'Quicksand';
const C = {
  titleBg: '#C9A0C2', titleText: '#FFFFFF',
  headerBg: '#B9A4D1', headerText: '#FFFFFF',
  legendText: '#9C7FA8',
  bandPink: '#FDF1F6', bandLavender: '#F3EEFB',
  attendBg: '#E2F4EA', attendText: '#1B4332',
  skipBg: '#FBE3EC', skipText: '#B36A8C',
  emptyBg: '#FFFFFF',
  safeBg: '#E2F4EA', safeText: '#3F7D5E',
  overBg: '#F8CBD6', overText: '#9C3F57',
  border: '#E6D3E3',
  totalBg: '#D9C3E0', totalText: '#5B3D6B',
  daysOffHeaderBg: '#F7E8C9', daysOffText: '#8A6A2E',
  issuesHeaderBg: '#D8CCEB',
  issuesOk: '#E2F4EA', issuesWarn: '#FBEFC9', issuesInfo: '#E6E1F7', issuesErr: '#F8CBD6',
  quizBg: '#FFFBEA', quizText: '#7B5800',
  examBg: '#FFF0E6', examText: '#8B3500',
  lunchBg: '#FFF5E6', lunchText: '#A0622A'
};
const TAB_COLORS = {
  myCourses: '#F3B6CE', conflicts: '#C7AEDD', weekly: '#A8DCC3',
  daysOff: '#F4D58D', issues: '#C9C2DA', today: '#F2A9C4', attendance: '#B9D4E8'
};

function refreshTimetable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let issues = [];

  try {
    const masterSS = SpreadsheetApp.openByUrl(MASTER_URL);
    const master = masterSS.getSheetByName(MASTER_TAB);
    if (!master) throw new Error(`Tab "${MASTER_TAB}" not found in the master spreadsheet.`);
    const tz = Session.getScriptTimeZone();

    const lastRow = master.getLastRow();
    const startRow = findDataStartRow(master, lastRow);
    if (startRow === -1) throw new Error('Could not find any date rows in column A.');

    const fullData = master.getRange(startRow, 1, lastRow - startRow + 1, LAST_COL_TO_READ).getValues();
    const normMap = {};
    MY_COURSES.forEach(c => normMap[normalizeCourseCode(c)] = c);

    let sessions = [];
    fullData.forEach(row => {
      const date = parseDateFromCell(row[DATE_COL_INDEX]);
      if (!date) return;
      const room = String(row[ROOM_COL_INDEX] || '').trim();
      SLOT_COL_INDEXES.forEach((colIdx, slotIdx) => {
        const parsed = parseCourseCell(row[colIdx]);
        if (!parsed) return;
        let course = parsed.course;
        let sessionType = 'class';
        let examCode = null;

        if (MY_COURSES.indexOf(course) !== -1) {
          sessionType = 'class';
        } else if (EXAM_COURSE_MAP[course]) {
          const info = EXAM_COURSE_MAP[course];
          examCode = course;
          course = info.course;
          sessionType = info.type;
        } else {
          const normMatch = normMap[normalizeCourseCode(course)];
          if (normMatch) {
            issues.push({ type: 'INFO', detail: `Matched "${course}" to "${normMatch}" on ${formatDateKey(date, tz)} — spelling differed slightly.` });
            course = normMatch;
            sessionType = 'class';
          } else { return; }
        }

        sessions.push({ date, slotIndex: slotIdx, slotLabel: SLOT_LABELS[slotIdx], course, sessionNum: parsed.sessionNum, room, skipped: false, type: sessionType, examCode });
      });
    });
    sessions.sort((a, b) => a.date - b.date || a.slotIndex - b.slotIndex);

    if (sessions.length === 0) {
      issues.push({ type: 'ERROR', detail: 'No matching sessions found. Refresh ABORTED.' });
      logIssues(ss, issues); return;
    }

    const overrideMap = readAttendanceOverrides(ss, tz);

    // Totals and skip budget — class sessions only, exams excluded
    let totals = {}, skipBudget = {};
    MY_COURSES.forEach(c => totals[c] = sessions.filter(s => s.course === c && s.type === 'class').length);
    MY_COURSES.forEach(c => skipBudget[c] = Math.round(totals[c] * SKIP_BUDGET_PCT));
    MY_COURSES.forEach(c => { if (totals[c] === 0) issues.push({ type: 'WARNING', detail: `No regular sessions found for "${c}".` }); });

    const daysOffSheet = ss.getSheetByName(SHEET_DAYS_OFF);
    let daysOff = new Set();
    if (daysOffSheet && daysOffSheet.getLastRow() > 1) {
      daysOffSheet.getRange(2, 1, daysOffSheet.getLastRow() - 1, 1).getValues()
        .forEach(r => { if (r[0]) daysOff.add(formatDateKey(new Date(r[0]), tz)); });
    }

    // Groups for conflict detection — class sessions only
    let groups = {};
    sessions.filter(s => s.type === 'class').forEach(s => {
      const key = formatDateKey(s.date, tz) + '|' + s.slotIndex;
      (groups[key] = groups[key] || []).push(s);
    });

    let decisions = [];
    const sortedDateKeys = [...new Set(sessions.map(s => formatDateKey(s.date, tz)))].sort();

    sortedDateKeys.forEach(dateKey => {
      const slotKeysToday = Object.keys(groups).filter(k => k.startsWith(dateKey + '|'))
        .sort((a, b) => parseInt(a.split('|')[1]) - parseInt(b.split('|')[1]));

      if (daysOff.has(dateKey)) {
        let coursesToday = new Set();
        slotKeysToday.forEach(k => groups[k].forEach(s => {
          const ov = overrideMap[sessionKey(s, tz)];
          s.skipped = (ov === ATT_STATUS.ATTENDED) ? false : true;
          coursesToday.add(s.course);
        }));
        if (coursesToday.size) {
          decisions.push({ date: dateKey, slot: 'Full day', courses: [...coursesToday].join(' / '), attend: '— (Day off)', skip: [...coursesToday].map(c => '🚫 ' + c).join(' + '), reason: 'Personal day off — all classes skipped.' });
        }
        return;
      }

      slotKeysToday.forEach(k => {
        const group = groups[k];
        const attendedOnes = group.filter(s => overrideMap[sessionKey(s, tz)] === ATT_STATUS.ATTENDED);
        if (attendedOnes.length > 1) {
          issues.push({ type: 'WARNING', detail: `You marked ${attendedOnes.map(s => s.course).join(' and ')} both as Attended on ${dateKey} (${group[0].slotLabel}) — impossible clash. Please double check.` });
        }

        let undecided = [];
        group.forEach(s => {
          const ov = overrideMap[sessionKey(s, tz)];
          if (ov === ATT_STATUS.ATTENDED) { s.skipped = false; }
          else if (ov === ATT_STATUS.SKIPPED) { s.skipped = true; }
          else { undecided.push(s); }
        });

        if (group.length === 1) { if (undecided.length) undecided[0].skipped = false; return; }

        if (attendedOnes.length >= 1) {
          undecided.forEach(s => s.skipped = true);
          decisions.push({ date: dateKey, slot: group[0].slotLabel, courses: group.map(s => `${s.course} (${s.room})`).join(' / '), attend: `✅ ${attendedOnes[0].course} (${attendedOnes[0].room})`, skip: group.filter(s => s.skipped).map(s => `🚫 ${s.course}`).join(' + '), reason: 'You manually marked your actual attendance for this clash.' });
          return;
        }

        if (undecided.length === 0) {
          decisions.push({ date: dateKey, slot: group[0].slotLabel, courses: group.map(s => `${s.course} (${s.room})`).join(' / '), attend: '— (none)', skip: group.map(s => `🚫 ${s.course}`).join(' + '), reason: 'You manually marked all as Skipped.' });
          return;
        }

        if (undecided.length === 1) {
          undecided[0].skipped = false;
          decisions.push({ date: dateKey, slot: group[0].slotLabel, courses: group.map(s => `${s.course} (${s.room})`).join(' / '), attend: `✅ ${undecided[0].course} (${undecided[0].room})`, skip: group.filter(s => s.skipped).map(s => `🚫 ${s.course}`).join(' + '), reason: 'Only one option left after your manual entries.' });
          return;
        }

        const remaining = undecided.map(s => ({ s, rem: skipBudget[s.course] - sessions.filter(x => x.course === s.course && x.type === 'class' && x.skipped).length }));
        remaining.sort((a, b) => a.rem - b.rem);
        const keep = remaining[0];
        const drop = remaining.slice(1);
        drop.forEach(d => d.s.skipped = true);
        decisions.push({ date: dateKey, slot: group[0].slotLabel, courses: group.map(s => `${s.course} (${s.room})`).join(' / '), attend: `✅ ${keep.s.course} (${keep.s.room})`, skip: drop.map(d => `🚫 ${d.s.course}`).join(' + '), reason: `${keep.s.course} has the tightest skip budget (${keep.rem}). ` + drop.map(d => `${d.s.course} budget = ${skipBudget[d.s.course]}.`).join(' ') });
      });
    });

    let skipsUsed = {};
    MY_COURSES.forEach(c => skipsUsed[c] = sessions.filter(s => s.course === c && s.type === 'class' && s.skipped).length);

    writeMyCoursesGrid(ss, sessions, tz);
    writeTodayClasses(ss, sessions, tz);
    writeAttendanceTab(ss, sessions, tz, overrideMap);
    writeConflicts(ss, totals, skipBudget, skipsUsed, decisions, tz);
    writeWeeklySummary(ss, sessions, totals, skipBudget, skipsUsed, tz);
    styleDaysOffTab(ss);

    issues.push({ type: 'OK', detail: `Refreshed successfully. ${sessions.filter(s=>s.type==='class').length} classes + ${sessions.filter(s=>s.type!=='class').length} exams/quizzes found.` });
    logIssues(ss, issues);

  } catch (err) {
    logIssues(ss, [{ type: 'ERROR', detail: 'Refresh failed: ' + err.message + '. Existing tabs were left untouched.' }]);
    throw err;
  }
}

function normalizeCourseCode(s) { return String(s).trim().toUpperCase().replace(/\s+/g, ' '); }

function findDataStartRow(sheet, lastRow) {
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < colA.length; i++) { if (parseDateFromCell(colA[i][0])) return i + 1; }
  return -1;
}

function parseDateFromCell(cell) {
  if (cell instanceof Date) return cell;
  const text = String(cell).trim();
  if (!text) return null;
  let d = new Date(text);
  if (!isNaN(d)) return d;
  const m = text.match(/(\w+\s+\d{1,2},?\s+\d{4})/);
  if (m) { d = new Date(m[1]); if (!isNaN(d)) return d; }
  return null;
}

function parseCourseCell(cellValue) {
  const text = String(cellValue || '').trim();
  if (!text) return null;
  const m = text.match(/^(.+?)\s+(\d+)$/);
  if (m) return { course: m[1].trim(), sessionNum: parseInt(m[2], 10) };
  // No number — treat the whole cell as the course code (for exams/quizzes like Q1-ADMA, ET-TSB)
  if (EXAM_COURSE_MAP[text]) return { course: text, sessionNum: 0 };
  return null;
}

function formatDateKey(date, tz) { return Utilities.formatDate(date, tz, 'yyyy-MM-dd'); }
function formatDateDisplay(dateKey, tz) { return Utilities.formatDate(new Date(dateKey), tz, 'EEE, MMMM d, yyyy'); }
function sessionKey(s, tz) { return formatDateKey(s.date, tz) + '|' + s.slotIndex + '|' + s.course + '|' + s.sessionNum; }

// ====================== ✏️ Mark Attendance ======================
function readAttendanceOverrides(ss, tz) {
  const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  let map = {};
  if (sheet && sheet.getLastRow() > 3) {
    const numRows = sheet.getLastRow() - 3;
    const data = sheet.getRange(4, 1, numRows, 8).getValues();
    data.forEach(r => {
      const status = r[3];
      const dateVal = r[4], slotIdx = r[5], course = r[6], sessionNum = r[7];
      if (!dateVal || !course) return;
      const dateKey = formatDateKey(new Date(dateVal), tz);
      const key = dateKey + '|' + slotIdx + '|' + course + '|' + sessionNum;
      if (status === ATT_STATUS.ATTENDED || status === ATT_STATUS.SKIPPED) map[key] = status;
    });
  }
  return map;
}

function writeAttendanceTab(ss, sessions, tz, overrideMap) {
  const sheet = ss.getSheetByName(SHEET_ATTENDANCE) || ss.insertSheet(SHEET_ATTENDANCE);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  const todayKey = formatDateKey(new Date(), tz);
  // Exams/quizzes are compulsory — only regular classes need marking
  const relevant = sessions.filter(s => s.type === 'class' && formatDateKey(s.date, tz) <= todayKey)
    .slice().sort((a, b) => a.date - b.date || a.slotIndex - b.slotIndex);

  sheet.getRange(1, 1, 1, 4).merge().setValue('✏️  Mark What You Actually Attended  ✏️')
    .setBackground(C.titleBg).setFontColor(C.titleText).setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setFontFamily(FONT);
  sheet.getRange(2, 1, 1, 4).merge().setValue('Only today + past classes shown. Exams & quizzes are compulsory — not listed here.')
    .setFontColor(C.legendText).setFontStyle('italic').setFontSize(9)
    .setHorizontalAlignment('center').setFontFamily(FONT);

  if (!relevant.length) {
    sheet.getRange(4, 1, 1, 4).merge().setValue('Nothing to mark yet — check back once classes start!')
      .setBackground(C.bandLavender).setFontColor(C.legendText).setHorizontalAlignment('center').setFontFamily(FONT);
    sheet.setTabColor(TAB_COLORS.attendance);
    return;
  }

  const colHeaderRow = 3;
  sheet.getRange(colHeaderRow, 1, 1, 4).setValues([['Date', 'Class', 'Room', 'Status']])
    .setBackground(C.headerBg).setFontColor(C.headerText).setFontWeight('bold').setFontFamily(FONT).setHorizontalAlignment('center');

  let row = colHeaderRow + 1;
  let lastDateKey = null;
  let bandToggle = 0;

  relevant.forEach(s => {
    const dateKey = formatDateKey(s.date, tz);
    if (dateKey !== lastDateKey) { bandToggle = 1 - bandToggle; lastDateKey = dateKey; }
    const status = overrideMap[sessionKey(s, tz)] || '';
    const dateLabel = Utilities.formatDate(s.date, tz, 'EEE, d MMM');
    const classLabel = `${s.slotLabel.split(' ')[0]} · ${s.course} #${s.sessionNum}`;

    sheet.getRange(row, 1, 1, 4).setValues([[dateLabel, classLabel, s.room, status]])
      .setBackground(bandToggle ? C.bandPink : C.bandLavender).setFontFamily(FONT);
    if (status === ATT_STATUS.ATTENDED) sheet.getRange(row, 4).setFontColor(C.attendText).setFontWeight('bold');
    if (status === ATT_STATUS.SKIPPED) sheet.getRange(row, 4).setFontColor(C.skipText).setFontWeight('bold');

    sheet.getRange(row, 5).setValue(s.date);
    sheet.getRange(row, 6).setValue(s.slotIndex);
    sheet.getRange(row, 7).setValue(s.course);
    sheet.getRange(row, 8).setValue(s.sessionNum);
    row++;
  });

  const statusRange = sheet.getRange(colHeaderRow + 1, 4, row - colHeaderRow - 1, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([ATT_STATUS.ATTENDED, ATT_STATUS.SKIPPED], true)
    .setAllowInvalid(true).build());

  sheet.getRange(colHeaderRow, 1, row - colHeaderRow, 4).setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.hideColumns(5, 4);
  sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 90); sheet.setColumnWidth(4, 110);
  sheet.setFrozenRows(colHeaderRow);
  sheet.setTabColor(TAB_COLORS.attendance);
}

// ====================== 📌 Today's Classes ======================
function writeTodayClasses(ss, sessions, tz) {
  const sheet = ss.getSheetByName(SHEET_TODAY) || ss.insertSheet(SHEET_TODAY, 0);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  const todayKey = formatDateKey(new Date(), tz);
  const todaySessions = sessions.filter(s => formatDateKey(s.date, tz) === todayKey)
    .slice().sort((a, b) => a.slotIndex - b.slotIndex);
  const todayDisplay = Utilities.formatDate(new Date(), tz, 'EEEE, d MMMM yyyy');

  sheet.getRange(1, 1, 1, 4).merge().setValue(`📌  Today's Classes — ${todayDisplay}  📌`)
    .setBackground(C.titleBg).setFontColor(C.titleText).setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setFontFamily(FONT);

  if (!todaySessions.length) {
    sheet.getRange(3, 1, 1, 4).merge().setValue('🌸  No classes today — enjoy the day off!  🌸')
      .setBackground(C.attendBg).setFontColor(C.attendText).setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center').setFontFamily(FONT);
    sheet.setColumnWidths(1, 4, 180); sheet.setTabColor(TAB_COLORS.today);
    moveSheetFirst(ss, sheet); return;
  }

  const headerRow = 3;
  sheet.getRange(headerRow, 1, 1, 4).setValues([['Slot', 'Course', 'Room', 'Status']])
    .setBackground(C.headerBg).setFontColor(C.headerText).setFontWeight('bold').setFontFamily(FONT).setHorizontalAlignment('center');

  let row = headerRow + 1;
  todaySessions.forEach(s => {
    let bg, statusText, textColor;
    if (s.type === 'exam') {
      bg = C.examBg; statusText = '📋 COMPULSORY'; textColor = C.examText;
    } else if (s.type === 'quiz') {
      bg = C.quizBg; statusText = '📝 COMPULSORY'; textColor = C.quizText;
    } else {
      bg = s.skipped ? C.skipBg : C.attendBg;
      statusText = s.skipped ? '🚫 Skip' : '✓ Attend';
      textColor = s.skipped ? C.skipText : C.attendText;
    }

    const courseLabel = s.type !== 'class' ? `${s.examCode || s.course}` : `${s.course} #${s.sessionNum}`;
    sheet.getRange(row, 1, 1, 4).setValues([[s.slotLabel, courseLabel, s.room, statusText]])
      .setBackground(bg).setFontFamily(FONT);
    sheet.getRange(row, 4).setFontColor(textColor).setFontWeight('bold');
    if (s.type === 'class' && s.skipped) sheet.getRange(row, 2).setFontLine('line-through');
    row++;
  });

  sheet.getRange(headerRow, 1, row - headerRow, 4).setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidths(1, 1, 160); sheet.setColumnWidths(2, 1, 160);
  sheet.setColumnWidths(3, 1, 120); sheet.setColumnWidths(4, 1, 130);
  sheet.setFrozenRows(headerRow);
  sheet.setTabColor(TAB_COLORS.today);
  moveSheetFirst(ss, sheet);
}

function moveSheetFirst(ss, sheet) { ss.setActiveSheet(sheet); ss.moveActiveSheet(1); }

// ====================== 🌸 My Courses (grid with lunch column) ======================
function writeMyCoursesGrid(ss, sessions, tz) {
  const sheet = ss.getSheetByName(SHEET_MY_COURSES) || ss.insertSheet(SHEET_MY_COURSES);
  sheet.clear();
  sheet.clearConditionalFormatRules();
  if (!sessions.length) return;

  const numCols = 1 + DISPLAY_SLOTS.length; // 1 date col + 9 slots + 1 lunch = 11

  sheet.getRange(1, 1).setBackground(C.titleBg).setFontFamily(FONT);
  sheet.getRange(1, 2, 1, numCols - 1).merge().setValue('🌸  My Courses Timetable  🌸')
    .setBackground(C.titleBg).setFontColor(C.titleText).setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setFontFamily(FONT);

  sheet.getRange(2, 1).setBackground(C.bandLavender).setFontFamily(FONT);
  sheet.getRange(2, 2, 1, numCols - 1).merge()
    .setValue('✓ = attend  ·  🚫 = skip  ·  📝 = quiz (compulsory)  ·  📋 = end term (compulsory)')
    .setFontColor(C.legendText).setFontStyle('italic').setFontSize(9)
    .setHorizontalAlignment('center').setFontFamily(FONT);

  const headerRow = 3;
  sheet.getRange(headerRow, 1).setValue('Date');
  DISPLAY_SLOTS.forEach((slotIdx, dispIdx) => {
    const col = 2 + dispIdx;
    if (slotIdx === -1) {
      sheet.getRange(headerRow, col).setValue('🍽️ Lunch');
    } else {
      sheet.getRange(headerRow, col).setValue(SLOT_LABELS[slotIdx]);
    }
  });
  sheet.getRange(headerRow, 1, 1, numCols).setBackground(C.headerBg).setFontColor(C.headerText)
    .setFontWeight('bold').setHorizontalAlignment('center').setFontFamily(FONT);

  const minDate = new Date(Math.min(...sessions.map(s => s.date.getTime())));
  const maxDate = new Date(Math.max(...sessions.map(s => s.date.getTime())));

  let grid = {};
  sessions.forEach(s => {
    const key = formatDateKey(s.date, tz) + '|' + s.slotIndex;
    (grid[key] = grid[key] || []).push(s);
  });

  let dateLabels = [], richRows = [], bgRows = [];
  let cursor = new Date(minDate);
  let dayCount = 0;
  while (cursor <= maxDate) {
    const dateKey = formatDateKey(cursor, tz);
    dateLabels.push([Utilities.formatDate(cursor, tz, 'EEEE') + '\n' + Utilities.formatDate(cursor, tz, 'd MMM yyyy')]);
    let richRow = [], bgRow = [];
    DISPLAY_SLOTS.forEach(slotIdx => {
      if (slotIdx === -1) {
        richRow.push(SpreadsheetApp.newRichTextValue().setText('').build());
        bgRow.push(C.lunchBg);
      } else {
        const cellSessions = grid[dateKey + '|' + slotIdx] || [];
        const cell = buildSlotCell(cellSessions);
        richRow.push(cell.richText);
        bgRow.push(cell.background);
      }
    });
    richRows.push(richRow);
    bgRows.push(bgRow);
    cursor.setDate(cursor.getDate() + 1);
    dayCount++;
  }

  const dataStartRow = headerRow + 1;
  sheet.getRange(dataStartRow, 1, dayCount, 1).setValues(dateLabels)
    .setFontWeight('bold').setFontFamily(FONT).setVerticalAlignment('middle');
  sheet.getRange(dataStartRow, 2, dayCount, DISPLAY_SLOTS.length).setRichTextValues(richRows);
  sheet.getRange(dataStartRow, 2, dayCount, DISPLAY_SLOTS.length).setBackgrounds(bgRows);

  let dateBand = [];
  for (let i = 0; i < dayCount; i++) dateBand.push([i % 2 === 0 ? C.bandPink : C.bandLavender]);
  sheet.getRange(dataStartRow, 1, dayCount, 1).setBackgrounds(dateBand);

  sheet.getRange(headerRow, 1, dayCount + 1, numCols)
    .setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(dataStartRow, 2, dayCount, DISPLAY_SLOTS.length)
    .setFontFamily(FONT).setFontSize(9).setWrap(true).setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 115);
  DISPLAY_SLOTS.forEach((slotIdx, dispIdx) => {
    sheet.setColumnWidth(2 + dispIdx, slotIdx === -1 ? 70 : 125);
  });
  sheet.setRowHeights(dataStartRow, dayCount, 46);
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(1);
  sheet.setTabColor(TAB_COLORS.myCourses);
}

function buildSlotCell(cellSessions) {
  if (!cellSessions.length) {
    return { richText: SpreadsheetApp.newRichTextValue().setText('').build(), background: C.emptyBg };
  }

  const sorted = cellSessions.slice().sort((a, b) => {
    const order = s => s.type === 'exam' ? 0 : s.type === 'quiz' ? 1 : s.skipped ? 3 : 2;
    return order(a) - order(b);
  });

  const hasExam = sorted.some(s => s.type === 'exam');
  const hasQuiz = sorted.some(s => s.type === 'quiz');
  const classItems = sorted.filter(s => s.type === 'class');
  const allClassesSkipped = classItems.length > 0 && classItems.every(s => s.skipped);

  let bg;
  if (hasExam) bg = C.examBg;
  else if (hasQuiz) bg = C.quizBg;
  else if (allClassesSkipped) bg = C.skipBg;
  else bg = C.attendBg;

  let parts = [];
  sorted.forEach(s => {
    if (s.type === 'exam') {
      parts.push({ text: `📋 ${s.examCode}`, skipped: false, color: C.examText, bold: true });
      parts.push({ text: `(${s.room})`, skipped: false, color: C.examText, bold: false });
    } else if (s.type === 'quiz') {
      parts.push({ text: `📝 ${s.examCode}`, skipped: false, color: C.quizText, bold: true });
      parts.push({ text: `(${s.room})`, skipped: false, color: C.quizText, bold: false });
    } else {
      const marker = classItems.length > 1 ? (s.skipped ? '🚫 ' : '✓ ') : '';
      parts.push({ text: `${marker}${s.course} #${s.sessionNum}`, skipped: s.skipped, color: s.skipped ? C.skipText : C.attendText, bold: false });
      parts.push({ text: `(${s.room})`, skipped: s.skipped, color: s.skipped ? C.skipText : C.attendText, bold: false });
    }
  });

  const fullText = parts.map(p => p.text).join('\n');
  const builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  let pos = 0;
  parts.forEach(p => {
    const start = pos, end = pos + p.text.length;
    let style = SpreadsheetApp.newTextStyle().setForegroundColor(p.color);
    if (p.skipped) style = style.setStrikethrough(true);
    if (p.bold) style = style.setBold(true);
    builder.setTextStyle(start, end, style.build());
    pos = end + 1;
  });

  return { richText: builder.build(), background: bg };
}

// ====================== ⚠️ Conflicts ======================
function writeConflicts(ss, totals, skipBudget, skipsUsed, decisions, tz) {
  const sheet = ss.getSheetByName(SHEET_CONFLICTS) || ss.insertSheet(SHEET_CONFLICTS);
  sheet.clear(); sheet.clearConditionalFormatRules();
  let row = 1;

  sheet.getRange(row, 1, 1, 6).merge().setValue('⚠️  Conflict Resolutions — Smart Skip Plan  ⚠️')
    .setBackground(C.titleBg).setFontColor(C.titleText).setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setFontFamily(FONT);
  row += 2;

  sheet.getRange(row, 1).setValue('📊  Skip Budget Overview').setFontWeight('bold').setFontColor(C.legendText).setFontFamily(FONT); row++;
  sheet.getRange(row, 1, 1, 6).setValues([['Course', 'Total Classes', 'Skip Budget', 'Skips Used', 'Remaining', 'Status']])
    .setBackground(C.headerBg).setFontColor(C.headerText).setFontWeight('bold').setFontFamily(FONT).setHorizontalAlignment('center'); row++;
  const budgetStartRow = row;
  MY_COURSES.forEach((c, i) => {
    const remaining = skipBudget[c] - skipsUsed[c];
    sheet.getRange(row, 1, 1, 6).setValues([[c, totals[c], skipBudget[c], skipsUsed[c], remaining, remaining >= 0 ? '✅ Safe' : '❗ Over Budget']])
      .setBackground(i % 2 === 0 ? C.bandPink : C.bandLavender).setFontFamily(FONT);
    sheet.getRange(row, 6).setBackground(remaining >= 0 ? C.safeBg : C.overBg).setFontColor(remaining >= 0 ? C.safeText : C.overText).setFontWeight('bold');
    row++;
  });
  sheet.getRange(budgetStartRow - 1, 1, row - budgetStartRow + 1, 6).setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  row++;

  sheet.getRange(row, 1).setValue('📋  Conflict-by-Conflict Decisions').setFontWeight('bold').setFontColor(C.legendText).setFontFamily(FONT); row++;
  sheet.getRange(row, 1, 1, 6).setValues([['Date', 'Time Slot', 'Conflicting Classes', 'Attend', 'Skip', 'Why This Decision']])
    .setBackground(C.headerBg).setFontColor(C.headerText).setFontWeight('bold').setFontFamily(FONT).setHorizontalAlignment('center'); row++;
  const decisionStartRow = row;
  decisions.forEach((d, i) => {
    sheet.getRange(row, 1, 1, 6).setValues([[formatDateDisplay(d.date, tz), d.slot, d.courses, d.attend, d.skip, d.reason]])
      .setBackground(i % 2 === 0 ? C.bandPink : C.bandLavender).setFontFamily(FONT).setWrap(true).setVerticalAlignment('middle');
    sheet.getRange(row, 4).setFontColor(C.attendText).setFontWeight('bold');
    sheet.getRange(row, 5).setFontColor(C.skipText).setFontWeight('bold');
    row++;
  });
  if (decisions.length) sheet.getRange(decisionStartRow - 1, 1, row - decisionStartRow + 1, 6).setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidths(1, 1, 130); sheet.setColumnWidths(2, 1, 140);
  sheet.setColumnWidths(3, 2, 220); sheet.setColumnWidths(5, 1, 160); sheet.setColumnWidths(6, 1, 320);
  sheet.setFrozenRows(3);
  sheet.setTabColor(TAB_COLORS.conflicts);
}

// ====================== 📊 Weekly Summary ======================
function writeWeeklySummary(ss, sessions, totals, skipBudget, skipsUsed, tz) {
  const sheet = ss.getSheetByName(SHEET_WEEKLY) || ss.insertSheet(SHEET_WEEKLY);
  sheet.clear(); sheet.clearConditionalFormatRules();
  if (!sessions.length) return;

  const allDates = sessions.map(s => s.date.getTime());
  let weekStart = new Date(Math.min(...allDates));
  weekStart.setDate(weekStart.getDate() - ((weekStart.getDay() + 6) % 7));
  const maxDate = new Date(Math.max(...allDates));

  let weeks = [];
  let cursor = new Date(weekStart);
  while (cursor <= maxDate) { weeks.push(new Date(cursor)); cursor.setDate(cursor.getDate() + 7); }

  let header = ['Subject', 'Skip Budget', 'Total Classes'];
  weeks.forEach(w => header.push(Utilities.formatDate(w, tz, 'd MMM')));
  header.push('Classes to Attend', 'Skips Used', 'Skips Left');

  sheet.getRange(1, 1).setBackground(C.titleBg).setFontFamily(FONT);
  sheet.getRange(1, 2, 1, header.length - 1).merge().setValue('📊  Weekly Summary — Classes to Attend per Subject  📊')
    .setBackground(C.titleBg).setFontColor(C.titleText).setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setFontFamily(FONT);

  sheet.getRange(2, 1, 1, header.length).setValues([header])
    .setBackground(C.headerBg).setFontColor(C.headerText).setFontWeight('bold').setFontFamily(FONT).setHorizontalAlignment('center');

  let totalsRow = ['TOTAL', 0, 0, ...weeks.map(() => 0), 0, 0, 0];
  const weeklyColStart = 4;

  MY_COURSES.forEach((c, idx) => {
    const r = 3 + idx;
    let rowData = [c, skipBudget[c], totals[c]];
    weeks.forEach((w, wi) => {
      const wEnd = new Date(w); wEnd.setDate(wEnd.getDate() + 6);
      const count = sessions.filter(s => s.course === c && s.type === 'class' && !s.skipped && s.date >= w && s.date <= wEnd).length;
      rowData.push(count || '');
      totalsRow[3 + wi] += count;
    });
    const attend = totals[c] - skipsUsed[c];
    rowData.push(attend, skipsUsed[c], skipBudget[c] - skipsUsed[c]);
    sheet.getRange(r, 1, 1, header.length).setValues([rowData])
      .setBackground(idx % 2 === 0 ? C.bandPink : C.bandLavender).setFontFamily(FONT);
    sheet.getRange(r, 1).setFontWeight('bold');
    totalsRow[1] += skipBudget[c]; totalsRow[2] += totals[c];
    totalsRow[header.length - 3] += attend; totalsRow[header.length - 2] += skipsUsed[c]; totalsRow[header.length - 1] += (skipBudget[c] - skipsUsed[c]);
  });

  const totalRowNum = 3 + MY_COURSES.length;
  sheet.getRange(totalRowNum, 1, 1, header.length).setValues([totalsRow])
    .setBackground(C.totalBg).setFontColor(C.totalText).setFontWeight('bold').setFontFamily(FONT);
  sheet.getRange(2, 1, totalRowNum - 1, header.length).setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);

  if (weeks.length) {
    const heatRange = sheet.getRange(3, weeklyColStart, MY_COURSES.length, weeks.length);
    sheet.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint(C.bandPink)
      .setGradientMidpointWithValue('#F0D9E8', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
      .setGradientMaxpoint('#D98AB3')
      .setRanges([heatRange]).build()]);
  }

  sheet.setColumnWidth(1, 110);
  for (let i = 1; i < header.length; i++) sheet.setColumnWidth(1 + i, 85);
  sheet.setFrozenRows(2); sheet.setFrozenColumns(1);
  sheet.setTabColor(TAB_COLORS.weekly);
}

// ====================== 🌸 Personal Days Off ======================
function styleDaysOffTab(ss) {
  const sheet = ss.getSheetByName(SHEET_DAYS_OFF);
  if (!sheet) return;
  sheet.getRange(1, 1).setBackground(C.daysOffHeaderBg).setFontColor(C.daysOffText).setFontWeight('bold').setFontFamily(FONT);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow > 1) {
    for (let r = 2; r <= lastRow; r++) {
      sheet.getRange(r, 1).setBackground(r % 2 === 0 ? C.bandLavender : C.bandPink).setFontFamily(FONT);
    }
  }
  sheet.setColumnWidth(1, 150);
  sheet.setTabColor(TAB_COLORS.daysOff);
}

// ====================== 🩺 Issues Log ======================
function logIssues(ss, issues) {
  const sheet = ss.getSheetByName(SHEET_ISSUES) || ss.insertSheet(SHEET_ISSUES);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Type', 'Detail']])
      .setBackground(C.issuesHeaderBg).setFontColor('#4A3B63').setFontWeight('bold').setFontFamily(FONT);
    sheet.setColumnWidth(1, 130); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 500);
    sheet.setFrozenRows(1); sheet.setTabColor(TAB_COLORS.issues);
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy, h:mm a');
  const rows = issues.length ? issues.map(i => [now, i.type, i.detail]) : [[now, 'OK', 'Refreshed with no issues.']];
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows).setFontFamily(FONT).setWrap(true);
  rows.forEach((r, i) => {
    const bg = r[1] === 'OK' ? C.issuesOk : r[1] === 'WARNING' ? C.issuesWarn : r[1] === 'INFO' ? C.issuesInfo : C.issuesErr;
    sheet.getRange(startRow + i, 1, 1, 3).setBackground(bg);
  });
  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > 200) sheet.deleteRows(2, dataRows - 200);
}