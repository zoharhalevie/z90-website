/**
 * Z90 Method — Lead Pipeline (Google Apps Script)
 * ------------------------------------------------
 * מה הקוד עושה:
 *  1. setupSheet()  — מריצים פעם אחת. יוצר את 5 הלשוניות (Leads / Clients / Cohorts / Proof / KPIs)
 *                     עם כותרות, רשימות סטטוס, ונוסחאות.
 *  2. doPost(e)     — נקודת הקצה של הטפסים. מקבל JSON, כותב שורה ב-Leads, שולח לך מייל,
 *                     ומחזיר תשובה אמיתית (כדי שהטופס ידע אם באמת נשלח).
 *                     תאימות לאחור: אם מגיע pdfBase64 (מה-intake הישן) — מצרף למייל.
 *  3. doGet()       — בדיקת חיים: פותחים את ה-URL בדפדפן ורואים {"ok":true}.
 *
 * הגדרות פריסה (Deploy → New deployment → Web app):
 *   Execute as: Me   |   Who has access: Anyone
 */

// ─────────────────────────── הגדרות ───────────────────────────
var NOTIFY_EMAIL = 'zoharhalevie@gmail.com';   // לאן נשלחת התראת ליד
var SHEET_ID     = '1_KG9VHGWsRn5S7mja9haUrhdcf5mL7I7pBgGFGSDaqo';                          // ריק = הגיליון שהסקריפט מחובר אליו (Extensions → Apps Script מתוך הגיליון)
var TZ           = 'Asia/Jerusalem';

var LEAD_STATUSES = ['ליד חדש','בקשר','שיחה נקבעה','שיחה התקיימה','חושב','אבחון נקבע','אבחון בוצע','לקוח','לא מתאים','לא הגיב','סירב'];
var CLIENT_STATUSES = ['Onboarding','פעיל','הקפאה','סיים','חודש','נשר'];
var CONCERNS = ['גב תחתון','נוקשות','חוסר כוח','פחד להיפצע','רוצה מבנה','אחר'];
var PRODUCTS = ['Z90','קורס','לא ברור'];

// ─────────────────────────── עזר ───────────────────────────
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  var ss = ss_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header_(sh, cols) {
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#F4CE14').setFontColor('#161616');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
}

function validation_(sh, colIndex, list, rows) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
  sh.getRange(2, colIndex, rows || 1000, 1).setDataValidation(rule);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────── הקמה חד-פעמית ───────────────────────────
function setupSheet() {
  // ── Leads ──
  var L = sheet_('Leads');
  header_(L, [
    'תאריך', 'lead_id', 'שם', 'טלפון', 'מה מפריע', 'מוצר', 'עמוד',
    'source', 'medium', 'campaign', 'content', 'referrer',
    'סטטוס', 'תגובה ראשונה (תאריך+שעה)', 'דקות לתגובה',
    'תאריך שיחה', 'הגיע לשיחה?', 'תוצאת שיחה',
    'תאריך אבחון', 'הגיע לאבחון?', 'החלטה', 'מסלול', 'סכום ₪', 'מקור מפורט (מהשיחה)', 'הערות'
  ]);
  validation_(L, 5, CONCERNS);
  validation_(L, 6, PRODUCTS);
  validation_(L, 13, LEAD_STATUSES);
  validation_(L, 17, ['כן','לא (no-show)']);
  validation_(L, 18, ['אבחון','ישיר','לא מתאים','חושב','לא נענה']);
  validation_(L, 20, ['כן','לא']);
  validation_(L, 21, ['Essential','Premium','Elite','קורס','לא המשיך']);
  // דקות לתגובה = (תגובה ראשונה − תאריך) × 1440
  L.getRange('O2:O1000').setFormula('=IF(AND(A2<>"",N2<>""),ROUND((N2-A2)*1440,0),"")');
  L.getRange('A2:A1000').setNumberFormat('dd/mm/yyyy hh:mm');
  L.getRange('N2:N1000').setNumberFormat('dd/mm/yyyy hh:mm');
  L.setColumnWidths(1, 25, 120);

  // ── Clients ──
  var C = sheet_('Clients');
  header_(C, [
    'client_id', 'שם', 'טלפון', 'מסלול', 'תאריך התחלה', 'יום 30', 'יום 60', 'יום 90',
    'סטטוס', 'מקור (lead_id)', 'סכום ₪', 'אמצעי תשלום',
    'טופס פתיחה ✓', 'הצהרת בריאות ✓', 'הסכם ✓',
    'מדד 1 — שם', 'בייסליין 1', 'שבוע 6 — 1', 'יום 90 — 1',
    'מדד 2 — שם', 'בייסליין 2', 'שבוע 6 — 2', 'יום 90 — 2',
    'ביטחון בגוף 1-10 (בייסליין)', 'ביטחון שבוע 6', 'ביטחון יום 90',
    'אימונים מתוכננים', 'אימונים בפועל', 'Adherence',
    'שיחת יום 60 ✓', 'החלטת המשך', 'עדות ✓ (קישור)', 'הפניה (שם)', 'הערות'
  ]);
  validation_(C, 4, ['Essential','Premium','Elite']);
  validation_(C, 9, CLIENT_STATUSES);
  validation_(C, 31, ['מחזור 2','Continue','Independent','סיים','טרם הוחלט']);
  C.getRange('F2:F1000').setFormula('=IF(E2<>"",E2+30,"")');
  C.getRange('G2:G1000').setFormula('=IF(E2<>"",E2+60,"")');
  C.getRange('H2:H1000').setFormula('=IF(E2<>"",E2+90,"")');
  C.getRange('AC2:AC1000').setFormula('=IF(AND(AA2<>"",AB2<>""),AB2/AA2,"")').setNumberFormat('0%');
  C.getRange('E2:H1000').setNumberFormat('dd/mm/yyyy');

  // ── Cohorts (קורס) ──
  var K = sheet_('Cohorts');
  header_(K, [
    'cohort_id', 'תאריך מפגש 1', 'שם', 'טלפון', 'מקור', 'שולם ₪',
    'ביטחון לבד 1-10 (לפני)', 'פחד 1-10 (לפני)',
    'נוכחות 1', 'נוכחות 2', 'נוכחות 3', 'נוכחות 4',
    'K1 טכניקה 6/8', 'K2 בוחר משקל', 'K3 מכונות 3/3', 'K4 חימום לבד', 'K5 סקיצה', 'K6 אימון עצמאי ×2', 'K7 סימני עצירה 3/3',
    'ביטחון לבד 1-10 (אחרי)', 'NPS 0-10', 'עדות ✓', 'צ׳ק-אין 30 יום: אימונים לבד', 'הערות'
  ]);
  for (var c = 9; c <= 19; c++) validation_(K, c, ['✓','✗']);

  // ── Proof ──
  var P = sheet_('Proof');
  header_(P, ['תאריך', 'לקוח / משתתף', 'סוג', 'הטקסט / קישור', 'אישור פרסום', 'איפה בשימוש', 'הערות']);
  validation_(P, 3, ['ציטוט','וידאו','מדד','מקרה','רגע מתהליך']);
  validation_(P, 5, ['שם פרטי','אנונימי','לא']);

  // ── KPIs (שבועי, נוסחאות) ──
  var W = sheet_('KPIs');
  header_(W, ['שבוע מתחיל ב-', 'לידים', 'לידים איכותיים*', 'חציון דקות לתגובה', 'שיחות נקבעו', 'שיחות התקיימו', 'Show rate', 'אבחונים', 'לקוחות חדשים', 'הכנסה ₪', 'הערות']);
  W.getRange('A2').setValue(new Date());
  for (var r = 2; r <= 27; r++) {
    if (r > 2) W.getRange(r, 1).setFormula('=A' + (r - 1) + '+7');
    W.getRange(r, 2).setFormula('=COUNTIFS(Leads!$A:$A,">="&A' + r + ',Leads!$A:$A,"<"&A' + r + '+7)');
    W.getRange(r, 3).setFormula('=COUNTIFS(Leads!$A:$A,">="&A' + r + ',Leads!$A:$A,"<"&A' + r + '+7,Leads!$M:$M,"<>לא מתאים",Leads!$M:$M,"<>לא הגיב",Leads!$M:$M,"<>ליד חדש")');
    W.getRange(r, 4).setFormula('=IFERROR(MEDIAN(FILTER(Leads!$O:$O,Leads!$A:$A>=A' + r + ',Leads!$A:$A<A' + r + '+7,Leads!$O:$O<>"")),"")');
    W.getRange(r, 5).setFormula('=COUNTIFS(Leads!$P:$P,">="&A' + r + ',Leads!$P:$P,"<"&A' + r + '+7)');
    W.getRange(r, 6).setFormula('=COUNTIFS(Leads!$P:$P,">="&A' + r + ',Leads!$P:$P,"<"&A' + r + '+7,Leads!$Q:$Q,"כן")');
    W.getRange(r, 7).setFormula('=IF(E' + r + '>0,F' + r + '/E' + r + ',"")');
    W.getRange(r, 8).setFormula('=COUNTIFS(Leads!$S:$S,">="&A' + r + ',Leads!$S:$S,"<"&A' + r + '+7,Leads!$T:$T,"כן")');
    W.getRange(r, 9).setFormula('=COUNTIFS(Clients!$E:$E,">="&A' + r + ',Clients!$E:$E,"<"&A' + r + '+7)');
    W.getRange(r, 10).setFormula('=SUMIFS(Clients!$K:$K,Clients!$E:$E,">="&A' + r + ',Clients!$E:$E,"<"&A' + r + '+7)');
  }
  W.getRange('A2:A27').setNumberFormat('dd/mm/yyyy');
  W.getRange('G2:G27').setNumberFormat('0%');
  W.getRange('A29').setValue('* ליד איכותי = כל ליד שיצא מ"ליד חדש" ולא סומן "לא מתאים"/"לא הגיב". סטטוסים מתעדכנים ידנית ב-Leads.');

  // מוחקים את Sheet1 הריק אם קיים
  var def = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('גיליון1');
  if (def && ss_().getSheets().length > 1) ss_().deleteSheet(def);

  Logger.log('הגיליון הוקם. עכשיו: Deploy → New deployment → Web app.');
}

// ─────────────────────────── קליטת ליד ───────────────────────────
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '{}';
    var d = {};
    try { d = JSON.parse(raw); } catch (err) { d = e.parameter || {}; }

    var name  = String(d.name  || '').trim();
    var phone = String(d.phone || '').trim();
    if (!name || !phone) return json_({ ok: false, error: 'missing name/phone' });

    var now = new Date();
    var leadId = 'L' + Utilities.formatDate(now, TZ, 'yyMMdd-HHmmss');

    // תאימות לאחור: intake הישן שולח pdfBase64 — לא כותבים אותו ל-Leads, רק מצרפים למייל
    var isIntake = !!d.pdfBase64 && !d.source && !d.concern;

    if (!isIntake) {
      var L = sheet_('Leads');
      // שורה פנויה ראשונה לפי עמודה A (לא appendRow — עמודת הנוסחאות O נחשבת "תוכן")
      var colA = L.getRange('A2:A').getValues();
      var row = 2;
      while (row - 2 < colA.length && colA[row - 2][0] !== '') row++;
      L.getRange(row, 1, 1, 13).setValues([[
        now, leadId, name, phone,
        d.concern || '', d.product || '', d.page || '',
        d.utm_source || '', d.utm_medium || '', d.utm_campaign || '', d.utm_content || '', d.referrer || '',
        'ליד חדש'
      ]]);
      if (d.notes) L.getRange(row, 25).setValue(d.notes);
    }

    // מייל אליך — טקסט קצר, בלי PDF ללידים
    var subject = isIntake
      ? 'שאלון פתיחה — ' + name
      : 'ליד: ' + name + ' · ' + (d.concern || '—') + ' · ' + (d.utm_source || d.page || 'ישיר');
    var body =
      'שם: ' + name + '\nטלפון: ' + phone +
      (d.concern ? '\nמה מפריע: ' + d.concern : '') +
      (d.product ? '\nמוצר: ' + d.product : '') +
      (d.page ? '\nעמוד: ' + d.page : '') +
      (d.utm_source ? '\nמקור: ' + d.utm_source + ' / ' + (d.utm_medium || '') + ' / ' + (d.utm_campaign || '') + ' / ' + (d.utm_content || '') : '') +
      (d.referrer ? '\nreferrer: ' + d.referrer : '') +
      '\n\nלענות בוואטסאפ: https://wa.me/972' + phone.replace(/\D/g, '').replace(/^0/, '') +
      '\nגיליון: ' + ss_().getUrl();

    var opts = { name: 'Z90 Leads' };
    if (d.pdfBase64) {
      opts.attachments = [Utilities.newBlob(Utilities.base64Decode(d.pdfBase64), 'application/pdf', d.filename || ('Z90 - ' + name + '.pdf'))];
    }
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body, opts);

    return json_({ ok: true, id: leadId });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return json_({ ok: true, service: 'z90-leads' }); }

// הגדרות חד-פעמיות: שם + אזור זמן (הורץ ב-16.9.2026)
function fixSettings() {
  var ss = ss_();
  ss.rename('Z90 CRM');
  ss.setSpreadsheetTimeZone('Asia/Jerusalem');
  ss.setSpreadsheetLocale('he_IL');
  Logger.log('done');
}
