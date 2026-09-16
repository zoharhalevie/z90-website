/* ============================================================
   Z90 — lead-form.js
   מחליף את כל קוד ה-leads הישן (leadPdfHTML / submitLead / sendLead / sendHero)
   ב-index.html, landing.html ו-yesodot-beshlita/index.html.
   מה השתנה:
     • אין html2pdf (הסר את תגית ה-<script src="...html2pdf..."> מה-<head>)
     • שולח JSON ומקבל תשובה אמיתית — "נשלח" מוצג רק אם השרת אישר
     • קורא UTM מה-URL, שומר ב-sessionStorage (נשאר גם אם עוברים עמוד), שולח עם הליד
     • שולח שאלת סינון (concern) + מוצר + שם עמוד + referrer
     • יורה אירוע lead_submit ל-GA4 ו-Lead ל-Meta Pixel (אם מותקנים)
   ============================================================ */
(function () {
  // ▼ החלף ב-URL של הפריסה החדשה של Apps Script (מסתיים ב-/exec)
  var LEAD_ENDPOINT = 'https://script.google.com/macros/s/AKfycbzFoctjRyOxXtNAJ9gdDLuONlNQ9JAdqAQLnPsb2Lizq2AMxT4P5kt3a1w01SlveEMN/exec';
  // ▼ שם העמוד — שנה בכל קובץ: 'index' / 'landing' / 'yesodot'
  var PAGE = document.body.getAttribute('data-page') || 'index';
  // ▼ מוצר ברירת מחדל לעמוד: 'Z90' או 'קורס'
  var PRODUCT_DEFAULT = document.body.getAttribute('data-product') || 'Z90';

  // ── UTM: קריאה מה-URL, שמירה לסשן ──
  var UTM_KEYS = ['utm_source', 'utm_medium', 'utm_campaign', 'utm_content'];
  function captureUtm() {
    try {
      var q = new URLSearchParams(location.search), found = false, out = {};
      UTM_KEYS.forEach(function (k) { var v = q.get(k); if (v) { out[k] = v; found = true; } });
      if (found) sessionStorage.setItem('z90_utm', JSON.stringify(out));
      if (!sessionStorage.getItem('z90_ref') && document.referrer && document.referrer.indexOf(location.host) === -1) {
        sessionStorage.setItem('z90_ref', document.referrer);
      }
    } catch (e) {}
  }
  function getUtm() { try { return JSON.parse(sessionStorage.getItem('z90_utm') || '{}'); } catch (e) { return {}; } }
  function getRef() { try { return sessionStorage.getItem('z90_ref') || ''; } catch (e) { return ''; } }
  captureUtm();

  // ── מסלול שנבחר בכפתורי המסלולים (index) ──
  window.chosenPlan = window.chosenPlan || '';
  window.pickPlan = function (plan) {
    window.chosenPlan = plan;
    var c = document.getElementById('contact') || document.getElementById('lead');
    if (c) c.scrollIntoView({ behavior: 'smooth' });
    setTimeout(function () { var f = document.getElementById('lname'); if (f) f.focus({ preventScroll: true }); }, 900);
  };

  // ── שליחה ──
  function track() {
    try { if (window.gtag) gtag('event', 'lead_submit', { page: PAGE, product: PRODUCT_DEFAULT }); } catch (e) {}
    try { if (window.fbq) fbq('track', 'Lead', { content_name: PRODUCT_DEFAULT }); } catch (e) {}
  }

  async function submitLead(ids, formEl, okEl, btn) {
    var name = (document.getElementById(ids.name) || {}).value || '';
    var phone = (document.getElementById(ids.phone) || {}).value || '';
    var concernEl = document.getElementById(ids.concern);
    var concern = concernEl ? concernEl.value : '';
    name = name.trim(); phone = phone.trim();

    if (!name || !phone) { alert('אנא מלא/י שם וטלפון.'); return false; }
    if (!/^0?5\d[\d\- ]{7,9}$/.test(phone) && !/^\+?972/.test(phone)) { alert('מספר הטלפון לא נראה תקין.'); return false; }
    if (concernEl && !concern) { alert('בחר/י מה הכי מפריע — זה עוזר לי להגיע לשיחה מוכן.'); return false; }

    var span = btn ? btn.querySelector('span') : null;
    var orig = span ? span.textContent : '';
    if (btn) { btn.disabled = true; if (span) span.textContent = 'שולח...'; }

    var utm = getUtm();
    var payload = {
      name: name, phone: phone, concern: concern,
      product: window.chosenPlan || PRODUCT_DEFAULT, page: PAGE,
      utm_source: utm.utm_source || '', utm_medium: utm.utm_medium || '',
      utm_campaign: utm.utm_campaign || '', utm_content: utm.utm_content || '',
      referrer: getRef()
    };

    try {
      var res = await fetch(LEAD_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // text/plain = בלי preflight
        body: JSON.stringify(payload),
        redirect: 'follow'
      });
      var data = await res.json();
      if (!data.ok) throw new Error(data.error || 'server');
      formEl.style.display = 'none'; okEl.style.display = 'block';
      track();
      return true;
    } catch (e) {
      if (btn) { btn.disabled = false; if (span) span.textContent = orig; }
      // נפילה חכמה: פותחים וואטסאפ עם הפרטים כבר בהודעה — הליד לא הולך לאיבוד
      var msg = 'היי זהר, השארתי פרטים באתר אבל הטופס לא נשלח. ' + name + ', ' + phone + (concern ? ', ' + concern : '');
      if (confirm('השליחה נכשלה. לפתוח וואטסאפ עם הפרטים?')) {
        window.open('https://wa.me/972546349923?text=' + encodeURIComponent(msg), '_blank');
      }
      return false;
    }
  }

  // ── פונקציות שהכפתורים הקיימים קוראים להן ──
  window.sendLead = function () {
    submitLead({ name: 'lname', phone: 'lphone', concern: 'lconcern' },
      document.getElementById('leadForm'),
      document.getElementById('sentOk') || document.getElementById('lcSent'),
      document.querySelector('#leadForm .send-btn'));
  };
  window.sendHero = function () {
    submitLead({ name: 'hname', phone: 'hphone', concern: 'hconcern' },
      document.getElementById('heroForm'),
      document.getElementById('heroSentOk'),
      document.querySelector('#heroForm .send-btn'));
  };
})();
