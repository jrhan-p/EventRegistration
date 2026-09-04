function doGet(e) {
  syncWebAppUrl_();
  const view = cleanText_(e.parameter.view || 'scanner');
  if (view === 'receipt') return renderReceipt_(e.parameter.token, e.parameter.event);
  if (view === 'admin') {
    const admin = HtmlService.createTemplateFromFile('AdminPage');
    // Unguarded, this throws before rendering whenever the master spreadsheet
    // is missing or its Settings tab was renamed — the exact state the page's
    // own "not set up yet" message exists to explain.
    let url = '';
    try { url = setting_('WEB_APP_URL'); } catch (err) { url = cleanText_(ScriptApp.getService().getUrl()); }
    admin.webAppUrl = url;
    admin.consoleBuild = APP.version;
    return page_(admin.evaluate().setTitle('Event admin'));
  }
  // The scanner itself lives on GitHub Pages (top-level, so cameras work);
  // any old scanner URL pointing here forwards people to it.
  return page_(HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:60px auto;padding:0 18px;text-align:center">' +
    '<h2>The event scanner has moved</h2>' +
    '<p><a href="' + SCANNER_BASE + '" style="font-size:18px">Open the event scanner</a></p>' +
    '<p style="color:#667085">Sign in there with your station PIN.</p></div>'
  ).setTitle('Event scanner'));
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
    if (fn === 'scan') return out(processScan(p.value, p.pin));
    if (fn === 'phoneLookup') return out(phoneLookup(p.pin, p.phone));
    if (fn === 'resolvePin') {
      const station = resolvePin_(p.pin);
      return out(station ? {
        ok: true, type: station.type, stationName: station.stationName,
        eventId: station.event.eventId, eventName: station.event.name,
        location: station.location || '', infoUrl: station.infoUrl || ''
      } : { ok: false, message: 'Operator PIN is incorrect.' });
    }
    if (fn === 'eventInfo') {
      const info = eventById_(p.event);
      return out(info ? { ok: true, name: info.name, date: info.date, location: info.location, status: info.status }
        : { ok: false, message: 'Unknown event.' });
    }
    if (fn === 'addStation') return out(Object.assign({ ok: true }, adminAddStation(p.pin, p.event, p.name)));
    // Parameter names avoid "sid"/"token": Google's /exec front end hijacks
    // requests carrying both of those (its own reserved names) and returns an
    // error page before the script ever runs.
    if (fn === 'setTwilio') return out(adminSetTwilio(p.pin, p.twSid, p.twToken, p.twFrom));
    if (fn === 'testSms') return out(adminTestSms(p.pin, p.to));
    if (fn === 'setWalletConfig') return out(adminSetWalletConfig(p.pin, p.sa, p.appleUrl));
    if (fn === 'passData') {
      const token = cleanText_(p.token);
      const events = allEvents_();
      for (let i = 0; i < events.length; i++) {
        const found = findRegistration_(token, events[i]);
        if (!found) continue;
        const r = rowToRecord_(found.values, found.index);
        if (r.status !== APP.statuses.ACTIVE) break;
        return out({
          ok: true, fullName: r.fullName, registrationId: r.registrationId,
          eventId: events[i].eventId, eventName: events[i].name,
          eventDate: cleanText_(events[i].date), location: cleanText_(events[i].location),
          qr: APP.qrPrefix + token
        });
      }
      return out({ ok: false, message: 'Registration not found.' });
    }
    if (fn === 'listEvents') return out(Object.assign({ ok: true }, adminListEvents(p.pin)));
    if (fn === 'createEvent') return out(Object.assign({ ok: true }, adminCreateEvent(p.pin, p.name, p.date, p.location)));
    if (fn === 'eventAction') return out(Object.assign({ ok: true }, adminEventAction(p.pin, p.event, p.action)));
    // No admin PIN required: this tells the page whether to open in
    // first-run, locked, or recovery mode before anything is unlocked.
    if (fn === 'status') return out(adminStatus());
    // No admin PIN required: these two ARE the recovery path. The code goes
    // only to the organizer addresses on file, never to the caller.
    if (fn === 'requestPinReset') return out(adminRequestPinReset());
    if (fn === 'resetPin') return out(adminResetPinWithCode(p.code, p.newPin));
    if (fn === 'installMenu') return out(Object.assign({ ok: true }, adminInstallMenu(p.pin)));
    if (fn === 'setStationPin') return out(Object.assign({ ok: true }, adminSetStationPin(p.pin, p.event, p.station, p.newPin)));
    if (fn === 'setAdminPin') return out(adminChangeAdminPin(p.pin, p.newPin));
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

function renderReceipt_(token, eventId) {
  const clean = cleanText_(token);
  const preferred = eventById_(eventId);
  const candidates = preferred ? [preferred] : allEvents_();
  for (let i = 0; i < candidates.length; i++) {
    const found = findRegistration_(clean, candidates[i]);
    if (!found) continue;
    const event = candidates[i];
    const template = HtmlService.createTemplateFromFile('Receipt');
    const record = rowToRecord_(found.values, found.index);
    template.record = record;
    template.eventName = event.name;
    template.eventDate = event.date;
    template.location = event.location;
    template.qrUrl = 'https://quickchart.io/qr?size=320&text=' + encodeURIComponent(APP.qrPrefix + clean);
    template.walletButtons = walletButtons_(Object.assign({ token: clean }, record), event);
    return page_(template.evaluate().setTitle('Registration receipt'));
  }
  return page_(HtmlService.createHtmlOutput('<h2 style="font-family:Arial;text-align:center;margin-top:40px">Registration not found</h2>'));
}
