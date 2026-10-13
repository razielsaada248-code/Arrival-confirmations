// ===== הגדרות =====
const COUPLE_NAMES = 'אליה ואחינעם';
const FORM_URL = 'https://razielsaada248-code.github.io/Arrival-confirmations/'; // הקישור לטופס (GitHub Pages), משמש בהודעות התזכורת
// ==================

const SHEET_RESPONSES = 'תשובות';
const SHEET_SUMMARY = 'סיכום';
const SHEET_REMINDERS = 'תזכורות';

const HEADERS = ['תאריך תשובה', 'שם פרטי', 'שם משפחה', 'טלפון', 'צד', 'סטטוס', 'מספר מגיעים', 'שם המלווה', 'עודכן לאחרונה', 'היסטוריה'];
const COL = { date: 1, first: 2, last: 3, phone: 4, side: 5, status: 6, guests: 7, companion: 8, updated: 9, history: 10 };

const COMING = 'מגיע';
const NOT_COMING = 'לא מגיע';
const UNSURE = 'לא יודע עדיין';
const STATUSES = [COMING, NOT_COMING, UNSURE];
const SIDES = { groom: 'צד חתן', bride: 'צד כלה' };

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('אישורי הגעה')
    .addItem('התקנה ראשונית', 'setup')
    .addItem('רענון רשימת תזכורות', 'buildReminders')
    .addToUi();
}

function doGet() {
  return json({ status: 'ok', message: 'מערכת אישורי ההגעה פעילה' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website) return json({ status: 'ok' }); // מלכודת לבוטים

    const first = clean(d.firstName);
    const last = clean(d.lastName);
    const phone = normalizePhone(d.phone);
    const side = SIDES[d.side];
    const status = d.attendance;
    const guests = status === COMING ? Number(d.guests) : '';
    const companion = guests === 2 ? clean(d.companion) : '';

    if (!first || !last) return json({ status: 'error', message: 'חסר שם' });
    if (!phone) return json({ status: 'error', message: 'טלפון לא תקין' });
    if (!side) return json({ status: 'error', message: 'חסר צד' });
    if (STATUSES.indexOf(status) < 0) return json({ status: 'error', message: 'סטטוס לא תקין' });
    if (status === COMING && guests !== 1 && guests !== 2) return json({ status: 'error', message: 'מספר מגיעים לא תקין' });
    if (guests === 2 && !companion) return json({ status: 'error', message: 'חסר שם מלווה' });

    const sh = getSheet(SHEET_RESPONSES);
    const now = new Date();
    const row = findRowByPhone(sh, phone);

    if (row) {
      // אפשר לענות שוב רק מי שענה "לא יודע עדיין"
      const prev = sh.getRange(row, COL.status).getValue();
      if (prev !== UNSURE) return json({ status: 'already', previous: prev });
      const history = sh.getRange(row, COL.history).getValue();
      const entry = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM HH:mm') + ' ' + prev + ' → ' + status;
      sh.getRange(row, COL.first, 1, 9).setValues([[first, last, phone, side, status, guests, companion, now, history ? history + '\n' + entry : entry]]);
      buildReminders();
      return json({ status: 'updated' });
    }

    const newRow = sh.getLastRow() + 1;
    const range = sh.getRange(newRow, 1, 1, HEADERS.length);
    sh.getRange(newRow, COL.phone).setNumberFormat('@');
    range.setValues([[now, first, last, phone, side, status, guests, companion, now, '']]);
    if (status === UNSURE) buildReminders();
    return json({ status: 'ok' });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ===== התקנה: יוצר את שלוש הלשוניות, כותרות, נוסחאות וצבעים =====
function setup() {
  const ss = SpreadsheetApp.getActive();

  const resp = getSheet(SHEET_RESPONSES);
  resp.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#efe7d8');
  resp.setFrozenRows(1);
  resp.setRightToLeft(true);
  resp.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  resp.getRange('I:I').setNumberFormat('dd/MM/yyyy HH:mm');
  resp.getRange('D:D').setNumberFormat('@');
  resp.setColumnWidths(1, HEADERS.length, 130);
  resp.setColumnWidth(COL.history, 240);
  const statusRange = resp.getRange('F2:F');
  resp.setConditionalFormatRules([
    colorRule(statusRange, COMING, '#d9ead3'),
    colorRule(statusRange, NOT_COMING, '#f4cccc'),
    colorRule(statusRange, UNSURE, '#fff2cc')
  ]);

  buildSummary(getSheet(SHEET_SUMMARY));
  buildReminders();
  ss.setActiveSheet(ss.getSheetByName(SHEET_SUMMARY));
}

function buildSummary(sh) {
  sh.clear();
  sh.setRightToLeft(true);
  const R = "'" + SHEET_RESPONSES + "'!";
  const sideCond = (side) => (side ? ',' + R + 'E2:E,"' + side + '"' : '');
  const countStatus = (status, side) => '=COUNTIFS(' + R + 'F2:F,"' + status + '"' + sideCond(side) + ')';
  const columns = [null, SIDES.groom, SIDES.bride];

  const rows = [
    ['הזמנות שנשלחו (למלא ידנית בעמודות הצדדים)', () => '=C4+D4', () => 0],
    ['סה"כ ענו', (s) => (s ? '=COUNTIF(' + R + 'E2:E,"' + s + '")' : '=COUNTA(' + R + 'D2:D)')],
    ['מגיעים (מספר תשובות)', (s) => countStatus(COMING, s)],
    ['מגיעים (סה"כ אנשים)', (s) => '=SUMIFS(' + R + 'G2:G,' + R + 'F2:F,"' + COMING + '"' + sideCond(s) + ')'],
    ['לא מגיעים', (s) => countStatus(NOT_COMING, s)],
    ['עוד לא יודעים', (s) => countStatus(UNSURE, s)],
    ['לא ענו בכלל', (s, c) => '=MAX(0,' + c + '4-' + c + '5)'],
    ['אחוז שענו מתוך שנשלחו', (s, c) => '=IFERROR(' + c + '5/' + c + '4,"")'],
    ['אחוז מגיעים מתוך שנשלחו', (s, c) => '=IFERROR(' + c + '6/' + c + '4,"")']
  ];

  sh.getRange('A1').setValue('סיכום אישורי הגעה – ' + COUPLE_NAMES).setFontSize(16).setFontWeight('bold');
  sh.getRange('A3:D3').setValues([['נתון', 'סה"כ', SIDES.groom, SIDES.bride]]).setFontWeight('bold').setBackground('#efe7d8');

  rows.forEach((r, i) => {
    const rowNum = 4 + i;
    const values = [r[0]];
    columns.forEach((side, j) => {
      const letter = ['B', 'C', 'D'][j];
      const make = j > 0 && r[2] ? r[2] : r[1];
      values.push(make(side, letter));
    });
    sh.getRange(rowNum, 1, 1, 4).setValues([values]);
  });

  sh.getRange('C4:D4').setBackground('#fff2cc'); // תאים למילוי ידני
  sh.getRange('B11:D12').setNumberFormat('0%');
  sh.getRange('A7:D7').setFontWeight('bold').setBackground('#d9ead3');
  sh.setColumnWidth(1, 300);
  sh.setColumnWidths(2, 3, 110);
}

// ===== רשימת תזכורות: כל מי שענה "לא יודע עדיין", עם כפתור וואטסאפ =====
function buildReminders() {
  const sh = getSheet(SHEET_REMINDERS);
  sh.setRightToLeft(true);

  // שומר סימוני "נשלחה תזכורת" קיימים לפי טלפון
  const sent = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach((r) => {
      if (r[5] === true) sent[String(r[2])] = true;
    });
  }

  const resp = getSheet(SHEET_RESPONSES);
  const data = resp.getLastRow() > 1 ? resp.getRange(2, 1, resp.getLastRow() - 1, HEADERS.length).getValues() : [];
  const unsure = data.filter((r) => r[COL.status - 1] === UNSURE);

  sh.clear();
  sh.getRange(1, 1, 1, 6)
    .setValues([['שם פרטי', 'שם משפחה', 'טלפון', 'צד', 'שליחה', 'נשלחה תזכורת']])
    .setFontWeight('bold')
    .setBackground('#efe7d8');
  sh.setFrozenRows(1);
  if (!unsure.length) return;

  const rows = unsure.map((r) => {
    const first = r[COL.first - 1];
    const phone = String(r[COL.phone - 1]);
    const sideKey = r[COL.side - 1] === SIDES.bride ? 'bride' : 'groom';
    const link = FORM_URL ? FORM_URL + '?side=' + sideKey : '';
    const msg = 'היי ' + first + ', רצינו להזכיר לאשר הגעה לחתונה של ' + COUPLE_NAMES + (link ? ': ' + link : '') + ' 💛';
    const url = 'https://wa.me/972' + phone.slice(1) + '?text=' + encodeURIComponent(msg);
    return [first, r[COL.last - 1], phone, r[COL.side - 1], '=HYPERLINK("' + url + '","📲 שלח תזכורת")', !!sent[phone]];
  });

  sh.getRange(2, 3, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
  sh.getRange(2, 6, rows.length, 1).insertCheckboxes();
  rows.forEach((r, i) => sh.getRange(2 + i, 6).setValue(r[5]));
  sh.setColumnWidths(1, 6, 120);
}

// ===== עזר =====
function getSheet(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function findRowByPhone(sh, phone) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const phones = sh.getRange(2, COL.phone, last - 1, 1).getValues();
  for (let i = 0; i < phones.length; i++) {
    if (normalizePhone(phones[i][0]) === phone) return i + 2;
  }
  return 0;
}

function normalizePhone(raw) {
  let digits = String(raw || '').replace(/\D/g, '');
  if (digits.startsWith('972')) digits = '0' + digits.slice(3);
  if (/^5\d{8}$/.test(digits)) digits = '0' + digits; // אם גוגל הוריד את ה-0
  return /^0\d{8,9}$/.test(digits) ? digits : '';
}

function clean(s) {
  return String(s || '').trim().replace(/^[=+\-@]+/, '').slice(0, 60);
}

function colorRule(range, text, color) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
