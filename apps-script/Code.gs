// ===== הגדרות =====
const COUPLE_NAMES = 'אליה ואחינועם';
const FORM_URL = 'https://razielsaada248-code.github.io/Arrival-confirmations/'; // הקישור לטופס (GitHub Pages), משמש בהודעות התזכורת
const INVITATION_URL = FORM_URL + 'invitation.html'; // דף ההזמנה, נשלח בהודעה שבוע לפני
const EVENT_DATE = 'יום חמישי, אור לי\' כסלו התשפ"ז · 19.11.2026';
const EVENT_TIME = 'קבלת פנים 18:30 · חופה וקידושין 19:00';
const EVENT_VENUE = 'אולמי אטורה, פארק מיני ישראל לטרון';
const WAZE_URL = 'https://waze.com/ul?q=' + encodeURIComponent('אטורה מיני ישראל לטרון') + '&navigate=yes';
// מי שענה בטופס הצרפתי (fr/) מקבל את ההודעות בצרפתית
const COUPLE_NAMES_FR = 'Eliya et Ahinoam';
const FORM_URL_FR = FORM_URL + 'fr/';
const INVITATION_URL_FR = FORM_URL_FR + 'invitation.html';
const EVENT_DATE_FR = 'Jeudi 19 novembre 2026 (au soir du 10 Kislev 5787)';
const EVENT_TIME_FR = 'Réception 18h30 · Houppa et Kiddouchin 19h00';
const EVENT_VENUE_FR = 'Salles Atura, Parc Mini Israël, Latroun';
// טלפונים שאפשר לרשום מהם כמה מוזמנים שרוצים (כל שם = שורה נפרדת), ושלא נכנסים לרשימת "שבוע לפני"
const MULTI_GUEST_PHONES = ['0509303994'];
// נוסחי ה-SMS. {name} = שם פרטי, {link} = קישור. SMS בעברית: עד 70 תווים בחלק אחד, מעבר לזה כל 67 תווים = חלק נוסף
const SMS_REMINDER_HE = 'היי {name}, ראינו שעוד לא בטוח אם תגיעו לחתונת אליה ואחינועם. נשמח לתשובה סופית עד 9.11: {link}';
const SMS_REMINDER_FR = 'Bonjour {name}, vous n\'étiez pas encore sûr(e) de venir au mariage d\'Eliya et Ahinoam. Merci de répondre avant le 9.11 : {link}';
const SMS_EVENT_HE = 'היי {name}! מחכים לך בחתונת אליה ואחינועם, ה\' 19.11 18:30, אולמי אטורה לטרון. ההזמנה וניווט ב-Waze: {link}';
const SMS_EVENT_FR = 'Bonjour {name} ! Nous vous attendons au mariage d\'Eliya et Ahinoam, jeudi 19.11 à 18h30, salles Atura, Latroun. Faire-part et itinéraire : {link}';
// ==================

const SHEET_RESPONSES = 'תשובות';
const SHEET_SUMMARY = 'סיכום';
const SHEET_REMINDERS = 'תזכורות';
const SHEET_EVENT = 'שבוע לפני';

const HEADERS = ['תאריך תשובה', 'שם פרטי', 'שם משפחה', 'טלפון', 'צד', 'סטטוס', 'מספר מגיעים', 'שמות המלווים', 'עודכן לאחרונה', 'היסטוריה', 'מבוגרים', 'ילדים עד גיל 5', 'שפה', 'SMS תזכורת', 'SMS שבוע לפני'];
const COL = { date: 1, first: 2, last: 3, phone: 4, side: 5, status: 6, guests: 7, companion: 8, updated: 9, history: 10, adults: 11, kids: 12, lang: 13, smsReminder: 14, smsEvent: 15 };

const COMING = 'מגיע';
const NOT_COMING = 'לא מגיע';
const UNSURE = 'לא יודע עדיין';
const STATUSES = [COMING, NOT_COMING, UNSURE];
const SIDES = { groom: 'צד חתן', bride: 'צד כלה' };
const HEBREW = 'עברית';
const FRENCH = 'צרפתית';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('אישורי הגעה')
    .addItem('התקנה ראשונית', 'setup')
    .addItem('רענון רשימות השליחה', 'refreshLists')
    .addSeparator()
    .addItem('📱 שליחת תזכורת SMS', 'sendReminderSms')
    .addItem('📱 שליחת SMS שבוע לפני', 'sendEventSms')
    .addItem('SMS ניסיון למספר שלי', 'sendTestSms')
    .addItem('הגדרות חשבון SMS', 'setupSmsAccount')
    .addToUi();
}

// הזנה ידנית בלשונית "תשובות" מרעננת את רשימות השליחה
function onEdit(e) {
  if (e && e.range.getSheet().getName() === SHEET_RESPONSES) refreshLists();
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
    const lang = d.lang === 'fr' ? FRENCH : HEBREW;
    // צד החתן: 1 או 2 + שם המלווה. צד הכלה: מבוגרים + ילדים עד גיל 5.
    const bride = d.side === 'bride';
    const coming = status === COMING;
    const adults = coming ? (bride ? Number(d.adults) : Number(d.guests)) : '';
    const kids = coming ? (bride ? Number(d.kids) : 0) : '';
    const guests = coming ? adults + kids : '';
    const companion = !coming ? '' : bride ? clean(d.names, 300) : guests === 2 ? clean(d.companion) : '';

    if (!first || !last) return json({ status: 'error', message: 'חסר שם' });
    if (!phone) return json({ status: 'error', message: 'טלפון לא תקין' });
    if (!side) return json({ status: 'error', message: 'חסר צד' });
    if (STATUSES.indexOf(status) < 0) return json({ status: 'error', message: 'סטטוס לא תקין' });
    if (coming && bride && !(isCount(adults, 1, 10) && isCount(kids, 0, 10))) return json({ status: 'error', message: 'מספר מגיעים לא תקין' });
    if (coming && !bride && guests !== 1 && guests !== 2) return json({ status: 'error', message: 'מספר מגיעים לא תקין' });
    if (coming && !bride && guests === 2 && !companion) return json({ status: 'error', message: 'חסר שם מלווה' });

    const sh = getSheet(SHEET_RESPONSES);
    const now = new Date();
    // מטלפון "מרובה מוזמנים" מזהים לפי טלפון + שם מלא, כך שכל מוזמן חדש נרשם בשורה משלו
    const row = isMultiGuest(phone) ? findRowByPhoneAndName(sh, phone, first, last) : findRowByPhone(sh, phone);

    if (row) {
      // אפשר לענות שוב רק מי שענה "לא יודע עדיין"
      const prev = sh.getRange(row, COL.status).getValue();
      if (prev !== UNSURE) {
        // את הפרטים השמורים מחזירים רק אם גם השם הפרטי תואם, כדי שמספר טלפון לבד לא יחשוף תשובה של מישהו אחר
        const saved = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
        if (String(saved[COL.first - 1]).trim() !== first) return json({ status: 'already', previous: prev });
        const savedBride = saved[COL.side - 1] === SIDES.bride;
        return json({ status: 'already', previous: prev, saved: {
          firstName: String(saved[COL.first - 1]), lastName: String(saved[COL.last - 1]),
          phone: normalizePhone(saved[COL.phone - 1]) || phone,
          side: savedBride ? 'bride' : 'groom', attendance: prev,
          guests: String(saved[COL.guests - 1]), adults: String(saved[COL.adults - 1]), kids: String(saved[COL.kids - 1]),
          companion: savedBride ? '' : String(saved[COL.companion - 1]), names: savedBride ? String(saved[COL.companion - 1]) : ''
        } });
      }
      const history = sh.getRange(row, COL.history).getValue();
      const entry = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM HH:mm') + ' ' + prev + ' → ' + status;
      sh.getRange(row, COL.first, 1, 12).setValues([[first, last, phone, side, status, guests, companion, now, history ? history + '\n' + entry : entry, adults, kids, lang]]);
      refreshLists();
      return json({ status: 'updated' });
    }

    const newRow = sh.getLastRow() + 1;
    const range = sh.getRange(newRow, 1, 1, HEADERS.length);
    sh.getRange(newRow, COL.phone).setNumberFormat('@');
    range.setValues([[now, first, last, phone, side, status, guests, companion, now, '', adults, kids, lang, '', '']]);
    if (status !== NOT_COMING) refreshLists();
    return json({ status: 'ok' });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ===== התקנה: יוצר את הלשוניות, כותרות, נוסחאות וצבעים (אפשר להריץ שוב בבטחה) =====
function setup() {
  const ss = SpreadsheetApp.getActive();

  const resp = getSheet(SHEET_RESPONSES);
  resp.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#efe7d8');
  resp.setFrozenRows(1);
  resp.setRightToLeft(true);
  resp.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  resp.getRange('I:I').setNumberFormat('dd/MM/yyyy HH:mm');
  resp.getRange('N:O').setNumberFormat('dd/MM/yyyy HH:mm');
  resp.getRange('D:D').setNumberFormat('@');
  resp.setColumnWidths(1, HEADERS.length, 130);
  resp.setColumnWidth(COL.history, 240);
  const statusRange = resp.getRange('F2:F');
  resp.setConditionalFormatRules([
    colorRule(statusRange, COMING, '#d9ead3'),
    colorRule(statusRange, NOT_COMING, '#f4cccc'),
    colorRule(statusRange, UNSURE, '#fff2cc')
  ]);

  // רשימות בחירה להזנה ידנית
  const list = (values) => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  resp.getRange('E2:E').setDataValidation(list([SIDES.groom, SIDES.bride]));
  resp.getRange('F2:F').setDataValidation(list(STATUSES));
  resp.getRange('G2:G').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(1, 20).setAllowInvalid(false).setHelpText('סה"כ אנשים: צד חתן 1 או 2, צד כלה מבוגרים + ילדים').build()
  );
  const count = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 10).setAllowInvalid(false).build();
  resp.getRange('K2:L').setDataValidation(count);
  resp.getRange('M2:M').setDataValidation(list([HEBREW, FRENCH])); // ריק = עברית

  buildSummary(getSheet(SHEET_SUMMARY));
  buildSideSheets();
  refreshLists();
  protectSheets();
  ss.setActiveSheet(ss.getSheetByName(SHEET_SUMMARY));
}

// ===== הגנה: לשוניות שמחושבות לבד נעולות לעריכה, כדי ששינוי בטעות לא ישבור אותן =====
// בעל הגיליון תמיד יכול לערוך. את "תשובות" ואת רשימות השליחה לא נועלים, כי ממלאים בהן ידנית.
function protectSheets() {
  const summary = getSheet(SHEET_SUMMARY);
  lockSheet(summary, 'סיכום: מחושב לבד. רק התאים הצהובים (כמה הזמנות נשלחו) פתוחים לעריכה', [summary.getRange('C4:D4')]);
  Object.keys(SIDES).forEach((key) => {
    lockSheet(getSheet(SIDES[key]), SIDES[key] + ': מתעדכן לבד מתוך "תשובות". לשינוי תשובה עורכים בלשונית "תשובות"');
  });
}

function lockSheet(sh, description, editableRanges) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => p.remove());
  const p = sh.protect().setDescription(description);
  if (editableRanges) p.setUnprotectedRanges(editableRanges);
  const me = Session.getEffectiveUser();
  p.addEditor(me);
  p.removeEditors(p.getEditors().filter((u) => u.getEmail() !== me.getEmail()));
  if (p.canDomainEdit()) p.setDomainEdit(false);
}

function buildSummary(sh) {
  const sentCounts = sh.getRange('C4:D4').getValues()[0]; // שומר את מה שמולא ידנית
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
    ['אחוז מגיעים מתוך שנשלחו', (s, c) => '=IFERROR(' + c + '6/' + c + '4,"")'],
    ['מתוך המגיעים: ילדים עד גיל 5', (s) => '=SUMIFS(' + R + 'L2:L,' + R + 'F2:F,"' + COMING + '"' + sideCond(s) + ')'],
    ['מתוך המגיעים: מבוגרים', (s, c) => '=' + c + '7-' + c + '13']
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

  sh.getRange('C4:D4').setValues([sentCounts.map((v) => (typeof v === 'number' ? v : 0))]).setBackground('#fff2cc'); // תאים למילוי ידני
  sh.getRange('B11:D12').setNumberFormat('0%');
  sh.getRange('A7:D7').setFontWeight('bold').setBackground('#d9ead3');
  sh.setColumnWidth(1, 300);
  sh.setColumnWidths(2, 3, 110);
}

// ===== רשימות שליחה בוואטסאפ, עם כפתור לכל אחד =====
// "תזכורות": מי שענה "לא יודע עדיין" – תזכורת לאשר הגעה
// "שבוע לפני": מי שאישר הגעה – הודעה עם פרטי האירוע וההזמנה
function refreshLists() {
  buildReminders();
  buildEventReminders();
}

function isFrench(r) {
  return r[COL.lang - 1] === FRENCH;
}

function buildReminders() {
  buildList(SHEET_REMINDERS, UNSURE, '📲 שלח תזכורת', (r) => {
    if (isFrench(r)) return 'Bonjour ' + r[COL.first - 1] + ', nous voulions vous rappeler de confirmer votre présence au mariage d\'' + COUPLE_NAMES_FR + ' : ' + FORM_URL_FR + ' 💛';
    const sideKey = r[COL.side - 1] === SIDES.bride ? 'bride' : 'groom';
    return 'היי ' + r[COL.first - 1] + ', רצינו להזכיר לאשר הגעה לחתונה של ' + COUPLE_NAMES + ': ' + FORM_URL + '?side=' + sideKey + ' 💛';
  });
}

function buildEventReminders() {
  buildList(SHEET_EVENT, COMING, '📲 שלח הודעה', (r) => {
    const guests = Number(r[COL.guests - 1]) || 1;
    const kids = Number(r[COL.kids - 1]) || 0;
    if (isFrench(r)) {
      return [
        'Bonjour ' + r[COL.first - 1] + ' ! 💐',
        'Nous avons la joie de vous rappeler que dans une semaine nous fêterons le mariage d\'' + COUPLE_NAMES_FR + ' 🎉',
        '',
        '📅 ' + EVENT_DATE_FR,
        '🕡 ' + EVENT_TIME_FR,
        '📍 ' + EVENT_VENUE_FR,
        '🚗 Itinéraire : ' + WAZE_URL,
        '',
        'Vous avez confirmé la présence de ' + (guests === 1 ? 'une personne' : guests + ' personnes') + '.',
        '💌 Le faire-part : ' + INVITATION_URL_FR,
        '',
        'Au plaisir de vous voir !'
      ].join('\n');
    }
    const party = (guests === 1 ? 'אדם אחד' : guests + ' אנשים') + (kids ? ' (מתוכם ' + kids + ' ילדים)' : '');
    return [
      'היי ' + r[COL.first - 1] + '! 💐',
      'מזכירים בשמחה שבעוד שבוע נחגוג את החתונה של ' + COUPLE_NAMES + ' 🎉',
      '',
      '📅 ' + EVENT_DATE,
      '🕡 ' + EVENT_TIME,
      '📍 ' + EVENT_VENUE,
      '🚗 ניווט: ' + WAZE_URL,
      '',
      'אישרתם הגעה של ' + party + '.',
      '💌 ההזמנה: ' + INVITATION_URL,
      '',
      'מחכים לראותכם!'
    ].join('\n');
  }, (r) => isMultiGuest(r[COL.phone - 1]));
}

// skip: מי לא להכניס לרשימה (לא חובה)
function buildList(sheetName, status, buttonText, makeMessage, skip) {
  const sh = getSheet(sheetName);
  sh.setRightToLeft(true);

  // שומר סימוני "נשלח" קיימים לפי טלפון
  const sent = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach((r) => {
      if (r[5] === true) sent[String(r[2])] = true;
    });
  }

  const resp = getSheet(SHEET_RESPONSES);
  const data = resp.getLastRow() > 1 ? resp.getRange(2, 1, resp.getLastRow() - 1, HEADERS.length).getValues() : [];
  const people = data.filter((r) => r[COL.status - 1] === status && !(skip && skip(r)));

  sh.clear();
  sh.getRange(1, 1, 1, 7)
    .setValues([['שם פרטי', 'שם משפחה', 'טלפון', 'צד', 'שליחה', 'נשלח', 'SMS נשלח']])
    .setFontWeight('bold')
    .setBackground('#efe7d8');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 7, 120);
  if (!people.length) return;
  const smsCol = sheetName === SHEET_REMINDERS ? COL.smsReminder : COL.smsEvent;

  const rows = people.map((r) => {
    const phone = normalizePhone(r[COL.phone - 1]);
    const waNumber = phone.startsWith('+') ? phone.slice(1) : '972' + phone.slice(1);
    const url = 'https://wa.me/' + waNumber + '?text=' + encodeURIComponent(makeMessage(r));
    const button = phone ? '=HYPERLINK("' + url + '","' + buttonText + '")' : 'טלפון לא תקין';
    return [r[COL.first - 1], r[COL.last - 1], phone || r[COL.phone - 1], r[COL.side - 1], button, !!sent[phone], r[smsCol - 1]];
  });

  sh.getRange(2, 3, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  const checks = sh.getRange(2, 6, rows.length, 1);
  checks.insertCheckboxes();
  checks.setValues(rows.map((r) => [r[5]]));
}

// ===== לשונית נפרדת לכל צד: מתעדכנת לבד מתוך "תשובות" =====
function buildSideSheets() {
  const R = "'" + SHEET_RESPONSES + "'!";
  Object.keys(SIDES).forEach((key) => {
    const side = SIDES[key];
    const sh = getSheet(side);
    sh.clear();
    sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#efe7d8');
    sh.setFrozenRows(1);
    sh.getRange('A2').setFormula('=IFERROR(FILTER(' + R + 'A2:O,' + R + 'E2:E="' + side + '"),"אין עדיין תשובות")');
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange('I:I').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.setColumnWidths(1, HEADERS.length, 130);
    sh.setColumnWidth(COL.history, 240);
    const statusRange = sh.getRange('F2:F');
    sh.setConditionalFormatRules([
      colorRule(statusRange, COMING, '#d9ead3'),
      colorRule(statusRange, NOT_COMING, '#f4cccc'),
      colorRule(statusRange, UNSURE, '#fff2cc')
    ]);
  });
}

// ===== SMS בלחיצה אחת (דרך SMS4Free) =====
// פרטי החשבון נשמרים ב-Script Properties ולא בקוד, כי המאגר ציבורי. מזינים אותם דרך התפריט "הגדרות חשבון SMS".
const SMS_API_URL = 'https://api.sms4free.co.il/ApiSMS/v2/SendSMS';
const SMS_ERRORS = {
  '0': 'שגיאה כללית אצל SMS4Free', '-1': 'מפתח, שם משתמש או סיסמה שגויים', '-2': 'שם השולח לא תקין',
  '-3': 'מספר הטלפון לא תקין', '-4': 'נגמרו ההודעות בחבילה', '-5': 'תוכן ההודעה נחסם', '-6': 'מספר השולח עוד לא אומת'
};
const SMS_FATAL = ['-1', '-2', '-4', '-6']; // שגיאות שבגללן אין טעם להמשיך לשלוח לשאר

function setupSmsAccount() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const fields = [
    ['SMS4FREE_KEY', 'מפתח ה-API (באתר SMS4Free, בדף API)'],
    ['SMS4FREE_USER', 'שם המשתמש: מספר הנייד שאיתו נרשמת ל-SMS4Free'],
    ['SMS4FREE_PASS', 'הסיסמה של החשבון ב-SMS4Free'],
    ['SMS_SENDER', 'שם השולח: עד 11 אותיות באנגלית (למשל EliyaWedding), או מספר הנייד שלך. בלי חבילה בתשלום אפשר רק את מספר הנייד, ואם משאירים ריק נשלח ממנו']
  ];
  for (const [key, label] of fields) {
    const note = props.getProperty(key) ? '\n\n(כבר מוגדר. להשאיר ריק כדי לא לשנות)' : '';
    const res = ui.prompt('הגדרות חשבון SMS', label + note, ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const value = res.getResponseText().trim();
    if (value) props.setProperty(key, value);
  }
  ui.alert('נשמר. כדי לבדוק, אפשר לבחור "SMS ניסיון למספר שלי".');
}

function smsAccount() {
  const p = PropertiesService.getScriptProperties();
  const a = { key: p.getProperty('SMS4FREE_KEY'), user: p.getProperty('SMS4FREE_USER'), pass: p.getProperty('SMS4FREE_PASS') };
  a.sender = p.getProperty('SMS_SENDER') || a.user;
  return a.key && a.user && a.pass ? a : null;
}

// שולח SMS אחד. מחזיר { error, fatal }: error ריק אם הצליח
function sendSms(account, phone, text) {
  const res = UrlFetchApp.fetch(SMS_API_URL, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({ key: account.key, user: account.user, pass: account.pass, sender: account.sender, recipient: phone, msg: text })
  });
  let status;
  try {
    status = Number(JSON.parse(res.getContentText()).status);
  } catch (err) {
    return { error: 'תשובה לא צפויה מ-SMS4Free (' + res.getResponseCode() + ')', fatal: true };
  }
  if (status > 0) return { error: '' };
  return { error: SMS_ERRORS[String(status)] || 'שגיאה ' + status, fatal: SMS_FATAL.indexOf(String(status)) >= 0 };
}

function smsText(r, kind) {
  const fr = isFrench(r);
  let template, link;
  if (kind === 'reminder') {
    template = fr ? SMS_REMINDER_FR : SMS_REMINDER_HE;
    link = fr ? FORM_URL_FR : FORM_URL + '?side=' + (r[COL.side - 1] === SIDES.bride ? 'bride' : 'groom');
  } else {
    template = fr ? SMS_EVENT_FR : SMS_EVENT_HE;
    link = fr ? INVITATION_URL_FR : INVITATION_URL;
  }
  return template.replace('{name}', String(r[COL.first - 1]).trim()).replace('{link}', link);
}

// כמה הודעות זה יוריד מהחבילה: בעברית 70 תווים בחלק אחד (67 בכל חלק כשיש כמה), באנגלית/צרפתית 160 (153)
function smsParts(text) {
  const unicode = /[^\x00-\x7Féèàùç€]/.test(text);
  const one = unicode ? 70 : 160;
  const many = unicode ? 67 : 153;
  return text.length <= one ? 1 : Math.ceil(text.length / many);
}

// מי יקבל: לפי סטטוס, בלי מי שכבר קיבל. טלפון שמופיע בכמה שורות מקבל הודעה אחת
function smsRecipients(kind) {
  const status = kind === 'reminder' ? UNSURE : COMING;
  const sentCol = kind === 'reminder' ? COL.smsReminder : COL.smsEvent;
  const resp = getSheet(SHEET_RESPONSES);
  const data = resp.getLastRow() > 1 ? resp.getRange(2, 1, resp.getLastRow() - 1, HEADERS.length).getValues() : [];
  const byPhone = {};
  const skipped = [];
  data.forEach((r, i) => {
    if (r[COL.status - 1] !== status) return;
    if (kind === 'event' && isMultiGuest(r[COL.phone - 1])) return;
    const phone = normalizePhone(r[COL.phone - 1]);
    if (!phone || phone.startsWith('+')) {
      skipped.push(r[COL.first - 1] + ' ' + r[COL.last - 1]);
      return;
    }
    const g = byPhone[phone] || (byPhone[phone] = { phone: phone, row: r, rowNums: [], sent: false });
    g.rowNums.push(i + 2);
    if (r[sentCol - 1]) g.sent = true;
  });
  return { groups: Object.keys(byPhone).map((p) => byPhone[p]).filter((g) => !g.sent), skipped: skipped, sentCol: sentCol };
}

function sendReminderSms() { sendSmsBatch('reminder'); }
function sendEventSms() { sendSmsBatch('event'); }

function sendSmsBatch(kind) {
  const ui = SpreadsheetApp.getUi();
  const title = kind === 'reminder' ? 'תזכורת SMS' : 'SMS שבוע לפני';
  const account = smsAccount();
  if (!account) {
    ui.alert(title, 'קודם צריך להגדיר את חשבון ה-SMS: אישורי הגעה ← הגדרות חשבון SMS', ui.ButtonSet.OK);
    return;
  }
  const list = smsRecipients(kind);
  if (!list.groups.length) {
    ui.alert(title, 'אין למי לשלוח: כולם כבר קיבלו, או שאין אף אחד ברשימה.', ui.ButtonSet.OK);
    return;
  }
  const texts = list.groups.map((g) => smsText(g.row, kind));
  const parts = texts.reduce((n, t) => n + smsParts(t), 0);
  const answer = ui.alert(title,
    'יישלח ל-' + list.groups.length + ' אנשים (בערך ' + parts + ' הודעות מהחבילה).\n\nדוגמה:\n' + texts[0] +
    (list.skipped.length ? '\n\nלא יישלח ל-' + list.skipped.length + ' עם מספר מחו"ל או מספר לא תקין: ' + list.skipped.join(', ') : '') +
    '\n\nלשלוח עכשיו?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const resp = getSheet(SHEET_RESPONSES);
  const failed = [];
  let sent = 0;
  for (let i = 0; i < list.groups.length; i++) {
    const g = list.groups[i];
    const result = sendSms(account, g.phone, texts[i]);
    if (result.error) {
      failed.push(g.row[COL.first - 1] + ' ' + g.row[COL.last - 1] + ': ' + result.error);
      if (result.fatal) {
        failed.push('השליחה נעצרה, ' + (list.groups.length - i - 1) + ' נוספים לא נשלחו.');
        break;
      }
      continue;
    }
    const now = new Date();
    g.rowNums.forEach((n) => resp.getRange(n, list.sentCol).setValue(now)); // מסמן "נשלח SMS", כדי שלא יקבלו שוב
    sent++;
  }
  SpreadsheetApp.flush();
  refreshLists();
  ui.alert(title, 'נשלח ל-' + sent + ' אנשים.' + (failed.length ? '\n\nלא נשלח:\n' + failed.join('\n') : ''), ui.ButtonSet.OK);
}

// שולח את שני הנוסחים (תזכורת ושבוע לפני) למספר אחד, לבדיקה. לא מסמן כלום בגיליון
function sendTestSms() {
  const ui = SpreadsheetApp.getUi();
  const account = smsAccount();
  if (!account) {
    ui.alert('SMS ניסיון', 'קודם צריך להגדיר את חשבון ה-SMS: אישורי הגעה ← הגדרות חשבון SMS', ui.ButtonSet.OK);
    return;
  }
  const res = ui.prompt('SMS ניסיון', 'לאיזה מספר לשלוח? יישלחו 2 הודעות לדוגמה: תזכורת ושבוע לפני.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const phone = normalizePhone(res.getResponseText());
  if (!phone || phone.startsWith('+')) {
    ui.alert('SMS ניסיון', 'מספר לא תקין. צריך מספר ישראלי, למשל 0501234567.', ui.ButtonSet.OK);
    return;
  }
  const sample = [];
  sample[COL.first - 1] = 'בדיקה';
  sample[COL.side - 1] = SIDES.groom;
  sample[COL.lang - 1] = HEBREW;
  const errors = ['reminder', 'event'].map((kind) => sendSms(account, phone, smsText(sample, kind)).error).filter(Boolean);
  ui.alert('SMS ניסיון', errors.length ? 'לא נשלח: ' + errors.join(', ') : 'נשלחו 2 הודעות ל-' + phone + '.', ui.ButtonSet.OK);
}

// ===== עזר =====
function getSheet(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function isMultiGuest(phone) {
  return MULTI_GUEST_PHONES.indexOf(normalizePhone(phone)) >= 0;
}

function findRowByPhoneAndName(sh, phone, first, last) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const rows = sh.getRange(2, COL.first, lastRow - 1, 3).getValues(); // שם פרטי, שם משפחה, טלפון
  for (let i = 0; i < rows.length; i++) {
    if (normalizePhone(rows[i][2]) === phone && String(rows[i][0]).trim() === first && String(rows[i][1]).trim() === last) return i + 2;
  }
  return 0;
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

// מספר ישראלי נשמר כ-05X..., מספר מחו"ל (שהוזן עם + או 00) נשמר עם + וקידומת המדינה
function normalizePhone(raw) {
  const intl = /^\s*(\+|00)/.test(String(raw || ''));
  let digits = String(raw || '').replace(/\D/g, '');
  if (intl && digits.startsWith('00')) digits = digits.slice(2);
  if (digits.startsWith('972')) digits = '0' + digits.slice(3);
  else if (intl) return /^[1-9]\d{7,14}$/.test(digits) ? '+' + digits : '';
  if (/^5\d{8}$/.test(digits)) digits = '0' + digits; // אם גוגל הוריד את ה-0
  return /^0\d{8,9}$/.test(digits) ? digits : '';
}

function isCount(n, min, max) {
  return Number.isInteger(n) && n >= min && n <= max;
}

function clean(s, max) {
  return String(s || '').trim().replace(/^[=+\-@]+/, '').slice(0, max || 60);
}

function colorRule(range, text, color) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
