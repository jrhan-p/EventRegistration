function requireAdmin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH');
  if (!expected || hash_(String(pin)) !== expected) throw new Error('Admin PIN is incorrect.');
}

function setAdminPin(pin) {
  if (!/^\d{4,10}$/.test(String(pin))) throw new Error('The admin PIN must contain 4 to 10 digits.');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN_HASH', hash_(String(pin)));
}

function adminListEvents(pin) {
  requireAdmin_(pin);
  const sheet = db_().getSheetByName(SHEETS.EVENTS);
  const events = [];
  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getDataRange().getValues();
    const idx = headerIndex_(values[0]);
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const eventId = cleanText_(row[idx.event_id]);
      if (!eventId) continue;
      const lite = { eventId: eventId, spreadsheetId: cleanText_(row[idx.spreadsheet_id]) };
      let stations = [];
      try {
        ensureStations_(lite);
        stations = allStationRows_(lite).map(function(s) {
          return { id: s.id, name: s.name, type: s.type, pin: s.pin, location: s.location, active: s.active };
        });
      } catch (err) {}
      events.push({
        eventId: eventId,
        name: cleanText_(row[idx.event_name]),
        date: displayDate_(row[idx.event_date]),
        location: cleanText_(row[idx.location]),
        status: cleanText_(row[idx.status]).toUpperCase(),
        registrationUrl: cleanText_(row[idx.registration_url]),
        spreadsheetUrl: cleanText_(row[idx.spreadsheet_url]),
        formEditUrl: cleanText_(row[idx.form_edit_url]),
        stations: stations
      });
    }
  }
  const props = PropertiesService.getScriptProperties();
  return {
    masterUrl: db_().getUrl(), scannerUrl: SCANNER_BASE, events: events,
    build: APP.version,
    smsConfigured: Boolean(props.getProperty('TWILIO_ACCOUNT_SID') && props.getProperty('TWILIO_AUTH_TOKEN') && props.getProperty('TWILIO_FROM_NUMBER')),
    smsFrom: props.getProperty('TWILIO_FROM_NUMBER') || ''
  };
}

// Shown before anyone has unlocked anything, so it must be safe for a
// stranger to read: whether a PIN has been set at all, and a masked hint of
// where a reset code would go. The address itself is never returned, and a
// reset code is only ever mailed to it — never to whoever asked.
function adminStatus() {
  const out = { ok: true, build: APP.version, configured: false, ready: false, resetTo: '' };
  out.configured = Boolean(PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH'));
  try {
    out.resetTo = adminPinResetRecipients_().map(maskEmail_).filter(Boolean).join(', ');
    out.ready = true;
  } catch (err) {
    // No master spreadsheet yet — the very first bootstrap still needs the editor.
  }
  return out;
}

function maskEmail_(email) {
  const clean = cleanText_(email);
  const at = clean.indexOf('@');
  if (at < 1) return '';
  const name = clean.slice(0, at);
  return (name.length <= 2 ? name.charAt(0) : name.charAt(0) + '\u2026' + name.charAt(name.length - 1)) + clean.slice(at);
}

// Puts the RCCC Admin menu in the master spreadsheet without a trip to the
// editor. The trigger belongs to whoever the web app runs as, so the menu
// reaches people who also have script-project access — see Menu.gs.
function adminInstallMenu(pin) {
  requireAdmin_(pin);
  installAdminMenu_(null);
  return adminListEvents(pin);
}

function adminSetTwilio(pin, accountSid, authToken, fromNumber) {
  requireAdmin_(pin);
  const sid = cleanText_(accountSid);
  const token = cleanText_(authToken);
  const from = normalizePhone_(fromNumber);
  if (!/^AC[a-zA-Z0-9]{32}$/.test(sid)) throw new Error('That does not look like a Twilio Account SID (starts with AC, 34 characters).');
  if (!token) throw new Error('Enter the auth token.');
  if (!from) throw new Error('Enter the Twilio phone number, e.g. +18005551234.');
  setTwilioCredentials(sid, token, from);
  return { ok: true, smsConfigured: true, smsFrom: from };
}

function adminTestSms(pin, to) {
  requireAdmin_(pin);
  const phone = normalizePhone_(to);
  if (!phone) throw new Error('Enter a valid mobile number.');
  const result = sendSms_(phone, 'RCCC Events: test message — SMS is configured correctly.');
  if (result.skipped) throw new Error('Twilio is not configured yet — save the credentials first.');
  return { ok: true, sent: phone };
}

function adminCreateEvent(pin, name, date, location) {
  requireAdmin_(pin);
  if (!cleanText_(name)) throw new Error('Enter an event name.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    setupApplication();
    const sheet = db_().getSheetByName(SHEETS.EVENTS);
    const idx = headerIndex_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    const row = new Array(sheet.getLastColumn()).fill('');
    row[idx.event_name] = cleanText_(name);
    row[idx.event_date] = cleanText_(date);
    row[idx.location] = cleanText_(location);
    sheet.appendRow(row);
    provisionEvent_(sheet, idx, sheet.getLastRow(), row);
  } finally {
    lock.releaseLock();
  }
  return adminListEvents(pin);
}

function adminEventAction(pin, eventId, action) {
  requireAdmin_(pin);
  const wanted = cleanText_(action).toLowerCase();
  const cleanId = cleanText_(eventId);
  // A blank id would otherwise match any row whose event_id cell is empty —
  // exactly what a not-yet-provisioned row looks like.
  if (!cleanId) throw new Error('Unknown event.');
  // Locate AND mutate under the lock adminCreateEvent uses. deleteRow shifts
  // every row beneath it, so an unlocked delete could land on a row that moved
  // after it was located, or shift the row a 20-second provisionEvent_ is
  // still writing into — either way the registry ends up pointing at the
  // wrong spreadsheet and form.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Direct row scan (not eventById_) so half-provisioned ERROR rows — which
    // have no spreadsheet yet — can still be deleted.
    const sheet = db_().getSheetByName(SHEETS.EVENTS);
    const values = sheet.getDataRange().getValues();
    const idx = headerIndex_(values[0]);
    let rowNumber = 0;
    let row = null;
    for (let r = 1; r < values.length; r++) {
      if (cleanText_(values[r][idx.event_id]) !== cleanId) continue;
      // Never act on an ambiguous id: refuse rather than guess which row.
      if (rowNumber) throw new Error('Two rows in the Events sheet carry the id "' + cleanId + '". Give one of them a different id, then try again.');
      rowNumber = r + 1;
      row = values[r];
    }
    if (!rowNumber) throw new Error('Unknown event.');
    const formId = cleanText_(row[idx.form_id]);
    const spreadsheetId = cleanText_(row[idx.spreadsheet_id]);
    const status = cleanText_(row[idx.status]).toUpperCase();
    if (wanted === 'finalize') {
      if (!spreadsheetId) throw new Error('This event has no spreadsheet yet.');
      finalizeMealOrderFor_({ eventId: cleanId, spreadsheetId: spreadsheetId });
    } else if (wanted === 'close' || wanted === 'reopen') {
      if (!formId) throw new Error('This event has no form yet.');
      // Reopen is gated on STATUS, not on having ids. A row that failed part
      // way through provisioning has both ids recorded and still lacks its
      // form-submit trigger, so "has a spreadsheet id" is no proof it was ever
      // finished — reopening one would put a live form in front of guests that
      // nothing records: no ticket, no QR code, no email.
      if (wanted === 'reopen' && status !== 'CLOSED') {
        throw new Error('Only a closed event can be reopened — this one is ' + (status || 'not provisioned') +
          '. Delete it and create it again.');
      }
      FormApp.openById(formId).setAcceptingResponses(wanted === 'reopen');
      sheet.getRange(rowNumber, idx.status + 1).setValue(wanted === 'reopen' ? 'ACTIVE' : 'CLOSED');
    } else if (wanted === 'delete') {
      // Stop intake, detach the trigger, drop the registry row. Drive files
      // stay (no Drive scope) — the organizer can trash them from Drive.
      if (formId) {
        try { FormApp.openById(formId).setAcceptingResponses(false); } catch (err) {}
      }
      if (spreadsheetId) {
        detachFormTrigger_(spreadsheetId);
        // Retire the PINs now rather than leaving them live until the
        // resolvePin_ cache lapses five minutes from here.
        try {
          allStationRows_({ eventId: cleanId, spreadsheetId: spreadsheetId })
            .forEach(function(st) { forgetPin_(st.pin); });
        } catch (err) {}
      }
      sheet.deleteRow(rowNumber);
    } else {
      throw new Error('Unknown action "' + wanted + '".');
    }
  } finally {
    lock.releaseLock();
  }
  return adminListEvents(pin);
}

// Any station PIN can be reissued: a leaked PIN, a volunteer who left, or an
// organizer who just wants a memorable number. A blank newPin draws a fresh
// random one.
function adminSetStationPin(pin, eventId, stationId, newPin) {
  requireAdmin_(pin);
  const wantedStation = cleanText_(stationId);
  if (!wantedStation) throw new Error('Unknown station.');
  const clean = cleanText_(newPin);
  // Six digits is the floor because the station PIN is the only credential on
  // the scan, resolvePin, and phoneLookup endpoints, which are public and
  // unthrottled. Four digits would be ten thousand guesses.
  if (clean && !/^\d{6,8}$/.test(clean)) throw new Error('A station PIN must be 6 to 8 digits.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const event = eventById_(eventId);
    if (!event) throw new Error('Unknown event.');
    const sheet = ensureStations_(event);
    // Read every PIN in the system first: this opens each event's spreadsheet
    // and takes seconds, and the target row must not be located before it —
    // an organizer editing the Stations sheet meanwhile would shift it.
    const taken = allStationPins_(true);
    const values = sheet.getDataRange().getValues();
    const idx = headerIndex_(values[0]);
    let rowNumber = 0;
    let oldPin = '';
    for (let r = 1; r < values.length; r++) {
      if (cleanText_(values[r][idx.station_id]) !== wantedStation) continue;
      // Two rows sharing a station_id would leave the second one holding the
      // old PIN — still live, and unreachable from the console. Refuse.
      if (rowNumber) {
        throw new Error('Two rows in this event\'s Stations sheet use the id "' + wantedStation +
          '". Give one of them a different station_id, then try again.');
      }
      rowNumber = r + 1;
      oldPin = cleanText_(values[r][idx.pin]);
    }
    if (!rowNumber) throw new Error('Unknown station.');
    delete taken[oldPin];
    let assigned;
    if (!clean) {
      assigned = genStationPin_(taken);
    } else {
      if (taken[clean]) throw new Error('Another station already uses that PIN.');
      if (hash_(clean) === PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH')) {
        throw new Error('That is the admin PIN — give the station a different one.');
      }
      assigned = clean;
    }
    // Stored as text: Sheets would otherwise read 004321 back as 4321.
    sheet.getRange(rowNumber, idx.pin + 1).setNumberFormat('@').setValue(assigned);
    forgetPin_(oldPin);
    forgetPin_(assigned);
  } finally {
    lock.releaseLock();
  }
  return adminListEvents(pin);
}

// Changing the admin PIN needs the current one. Losing it is recovered from
// the master spreadsheet's RCCC Admin menu, where Google checks identity.
function adminChangeAdminPin(pin, newPin) {
  requireAdmin_(pin);
  applyNewAdminPin_(newPin);
  return { ok: true, adminPinChanged: true };
}

// Losing the admin PIN must not mean losing the system. The check here is
// control of an organizer mailbox — a real factor, already available, and one
// that does NOT require handing anybody the script project, which would also
// hand them the Twilio token and the Wallet signing key. The code is only ever
// mailed to addresses configured in the master spreadsheet's Settings sheet,
// never to an address supplied by the caller, so this public endpoint cannot
// be aimed at anyone.
function adminPinResetRecipients_() {
  const configured = setting_('ADMIN_EMAILS') || setting_('ORGANIZER_EMAIL');
  return cleanText_(configured).split(/[,;\s]+/).map(cleanText_).filter(Boolean);
}

function adminRequestPinReset() {
  const cache = CacheService.getScriptCache();
  // Generic answer either way: whether a mailbox is configured is not
  // something a stranger gets to learn, and the cooldown keeps the endpoint
  // from being used to flood the organizer (Gmail also caps daily sends).
  const generic = { ok: true, message: 'If an organizer address is on file, a reset code is on its way. It expires in 15 minutes.' };
  if (cache.get('pinreset:cooldown')) return generic;
  const to = adminPinResetRecipients_();
  if (!to.length) return generic;
  const code = String(Math.floor(Math.random() * 900000) + 100000);
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    ADMIN_RESET_HASH: hash_(code),
    ADMIN_RESET_EXPIRES: String(Date.now() + 15 * 60 * 1000),
    ADMIN_RESET_TRIES: '0'
  });
  cache.put('pinreset:cooldown', '1', 600);
  MailApp.sendEmail({
    to: to.join(','),
    subject: 'RCCC Events — admin PIN reset code: ' + code,
    body: 'Someone asked to reset the admin PIN for RCCC Community Events.\n\n' +
      'Reset code: ' + code + '\n\nIt expires in 15 minutes and works once.\n\n' +
      'Enter it in the admin console under "Admin PIN".\n\n' +
      'If this was not you, ignore this email — the current PIN still works and nothing has changed.'
  });
  return generic;
}

function adminResetPinWithCode(code, newPin) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('ADMIN_RESET_HASH');
  const expires = Number(props.getProperty('ADMIN_RESET_EXPIRES') || 0);
  const tries = Number(props.getProperty('ADMIN_RESET_TRIES') || 0);
  const clearReset = function() {
    props.deleteProperty('ADMIN_RESET_HASH');
    props.deleteProperty('ADMIN_RESET_EXPIRES');
    props.deleteProperty('ADMIN_RESET_TRIES');
  };
  if (!expected || !expires || Date.now() > expires) {
    clearReset();
    throw new Error('That code has expired. Request a new one.');
  }
  if (tries >= 5) {
    clearReset();
    throw new Error('Too many wrong codes. Request a new one.');
  }
  if (hash_(cleanText_(code)) !== expected) {
    props.setProperty('ADMIN_RESET_TRIES', String(tries + 1));
    throw new Error('That code is not right.');
  }
  // Validate the new PIN before burning the code, so a rejected PIN does not
  // force another round trip through the mailbox.
  applyNewAdminPin_(newPin);
  clearReset();
  return { ok: true, adminPinChanged: true };
}

function adminAddStation(pin, eventId, name) {
  requireAdmin_(pin);
  const event = eventById_(eventId);
  if (!event) throw new Error('Unknown event.');
  const clean = cleanText_(name);
  if (!clean) throw new Error('Enter a station name.');
  const sheet = ensureStations_(event);
  const stationId = 'S-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
  sheet.appendRow([stationId, clean, 'checkin', genStationPin_(allStationPins_(true)), true, '', '']);
  return adminListEvents(pin);
}

function adminCleanupLegacy(pin) {
  requireAdmin_(pin);
  cleanupLegacyData();
  return adminListEvents(pin);
}
