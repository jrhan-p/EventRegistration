function doGet(e) {
  syncWebAppUrl_();
  const view = cleanText_(e.parameter.view || 'scanner');
  if (view === 'receipt') return renderReceipt_(e.parameter.token, e.parameter.event);
  if (view === 'admin') {
    const admin = HtmlService.createTemplateFromFile('AdminPage');
    admin.webAppUrl = setting_('WEB_APP_URL');
    return page_(admin.evaluate().setTitle('Event admin'));
  }
  const event = eventById_(e.parameter.event);
  if (!event) return renderEventIndex_();
  const template = HtmlService.createTemplateFromFile('ScannerPage');
  template.mode = e.parameter.mode === APP.modes.MEAL ? APP.modes.MEAL : APP.modes.CHECKIN;
  template.eventName = event.name;
  template.eventId = event.eventId;
  template.webAppUrl = setting_('WEB_APP_URL');
  return page_(template.evaluate().setTitle('Event scanner'));
}

// Google's wrapper page only carries a mobile viewport tag when it is added
// here; a <meta> inside the sandboxed iframe does not scale the outer page.
function page_(output) {
  return output.addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTTP API used by the admin and scanner pages (plain fetch instead of
// google.script.run, which fails silently when browsers block third-party
// cookies) and testable directly with curl.
function doPost(e) {
  const out = function(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  };
  try {
    const p = (e && e.parameter) || {};
    const fn = cleanText_(p.fn);
    if (fn === 'scan') return out(processScan(p.value, p.pin, p.mode, p.event, p.meeting));
    if (fn === 'verifyPin') {
      const mode = p.mode === APP.modes.MEAL ? APP.modes.MEAL : APP.modes.CHECKIN;
      return out(authorizePin_(p.pin, mode) ? { ok: true } : { ok: false, message: 'Operator PIN is incorrect.' });
    }
    if (fn === 'eventInfo') {
      const info = eventById_(p.event);
      return out(info ? {
        ok: true, name: info.name, date: info.date, location: info.location, status: info.status,
        meetingName: meetingName_(info, p.meeting)
      } : { ok: false, message: 'Unknown event.' });
    }
    if (fn === 'addMeeting') return out(Object.assign({ ok: true }, adminAddMeeting(p.pin, p.event, p.name)));
    if (fn === 'listEvents') return out(Object.assign({ ok: true }, adminListEvents(p.pin)));
    if (fn === 'createEvent') return out(Object.assign({ ok: true }, adminCreateEvent(p.pin, p.name, p.date, p.location)));
    if (fn === 'eventAction') return out(Object.assign({ ok: true }, adminEventAction(p.pin, p.event, p.action)));
    if (fn === 'cleanupLegacy') return out(Object.assign({ ok: true }, adminCleanupLegacy(p.pin)));
    return out({ ok: false, message: 'Unknown function "' + fn + '".' });
  } catch (err) {
    return out({ ok: false, message: String((err && err.message) || err) });
  }
}

// Keeps the WEB_APP_URL setting matched to the live /exec deployment so emailed
// receipt links never go stale. Guarded against /dev-mode hits, which must not
// overwrite the public URL.
function syncWebAppUrl_() {
  try {
    const url = cleanText_(ScriptApp.getService().getUrl());
    if (url && /\/exec$/.test(url) && url !== setting_('WEB_APP_URL')) setSetting_('WEB_APP_URL', url);
  } catch (err) {
    console.warn('WEB_APP_URL sync skipped: ' + err);
  }
}

function renderEventIndex_() {
  const items = allEvents_().filter(function(event) { return event.status === 'ACTIVE'; })
    .map(function(event) {
      return '<li style="margin-bottom:10px"><b>' + html_(event.name) + '</b><br>' +
        '<a href="' + scannerUrl_(event.eventId, APP.modes.CHECKIN) + '">Check-in scanner</a> · ' +
        '<a href="' + scannerUrl_(event.eventId, APP.modes.MEAL) + '">Meal pickup scanner</a></li>';
    }).join('');
  const body = '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;padding:18px">' +
    '<h2>Active events</h2><ul>' + (items || '<li>No active events.</li>') + '</ul></div>';
  return page_(HtmlService.createHtmlOutput(body).setTitle('Event scanner'));
}

function renderReceipt_(token, eventId) {
  const clean = cleanText_(token);
  const preferred = eventById_(eventId);
  const candidates = preferred ? [preferred] : allEvents_();
  for (let i = 0; i < candidates.length; i++) {
    const found = findRegistration_(clean, candidates[i]);
    if (!found) continue;
    const event = candidates[i];
    const template = HtmlService.createTemplateFromFile('Receipt');
    template.record = rowToRecord_(found.values, found.index);
    template.eventName = event.name;
    template.eventDate = event.date;
    template.location = event.location;
    template.qrUrl = 'https://quickchart.io/qr?size=320&text=' + encodeURIComponent(APP.qrPrefix + clean);
    return page_(template.evaluate().setTitle('Registration receipt'));
  }
  return page_(HtmlService.createHtmlOutput('<h2 style="font-family:Arial;text-align:center;margin-top:40px">Registration not found</h2>'));
}
