function requireAdmin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH');
  // cleanText_ here too: every setter stores the hash of the trimmed PIN, so
  // hashing the raw parameter would reject a PIN with a stray space that the
  // page itself trimmed on the way in.
  if (!expected || hash_(cleanText_(pin)) !== expected) throw new Error('Admin PIN is incorrect.');
}

function setAdminPin(pin) {
  const clean = cleanText_(pin);
  if (!/^\d{6,10}$/.test(clean)) throw new Error('The admin PIN must be 6 to 10 digits.');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN_HASH', hash_(clean));
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
// hand them the Twilio token and the Wallet signing key.
//
// Both endpoints below are reachable by anyone who has the /exec URL, which is
// printed in every ticket email, so the design assumes a hostile caller:
//   - a wrong code never destroys the code sitting in the organizer's inbox,
//     because a stranger must not be able to hold recovery shut by guessing;
//   - a still-valid code is never replaced, so asking again cannot invalidate
//     what the organizer is holding;
//   - sends are capped per day and always leave headroom in the mail quota,
//     which is shared with the guests' ticket emails and is the scarcer
//     resource of the two;
//   - the answer is identical whatever happens, so nothing here reports
//     whether an address is configured or a mailbox is saturated.
// Guessing is bounded instead by the size of the code: 32^8, about a trillion.
const RESET_CODE_TTL_MS = 15 * 60 * 1000;
const RESET_MAX_PER_DAY = 8;
const RESET_MAIL_RESERVE = 25;
// No 0/O/1/I — this gets read off a screen and typed by hand.
const RESET_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

function adminPinResetRecipients_() {
  const configured = setting_('ADMIN_EMAILS') || setting_('ORGANIZER_EMAIL');
  // Filtered the same way here and in adminStatus, so the address the console
  // names is exactly the address the code is sent to.
  return cleanText_(configured).split(/[,;\s]+/).map(cleanText_)
    .filter(function(a) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a); });
}

function newResetCode_() {
  // getUuid is a random UUID; 256 % 32 === 0, so the mapping stays uniform.
  const hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/[^0-9a-f]/gi, '');
  let code = '';
  for (let i = 0; i < 8; i++) code += RESET_ALPHABET.charAt(parseInt(hex.substr(i * 2, 2), 16) % 32);
  return code;
}

function normalizeResetCode_(code) {
  return cleanText_(code).toUpperCase().replace(/[^0-9A-Z]/g, '');
}

function clearAdminReset_(props) {
  props.deleteProperty('ADMIN_RESET_HASH');
  props.deleteProperty('ADMIN_RESET_EXPIRES');
}

function adminRequestPinReset() {
  const generic = { ok: true, message: 'If an organizer address is on file, a reset code is on its way. It expires in 15 minutes.' };
  const lock = LockService.getScriptLock();
  // A burst collapses into one send instead of stacking: the check and the
  // send have to be one step, or every parallel caller passes the check.
  if (!lock.tryLock(10000)) return generic;
  try {
    const props = PropertiesService.getScriptProperties();
    // A code that is still good is never replaced — otherwise anyone could
    // invalidate the one the organizer is reading right now.
    if (Date.now() < Number(props.getProperty('ADMIN_RESET_EXPIRES') || 0)) return generic;
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const sentToday = props.getProperty('ADMIN_RESET_DAY') === today
      ? Number(props.getProperty('ADMIN_RESET_COUNT') || 0) : 0;
    if (sentToday >= RESET_MAX_PER_DAY) return generic;
    const to = adminPinResetRecipients_();
    if (!to.length) return generic;
    // Guests' tickets come out of this same daily allowance. Never spend the
    // last of it on a reset code.
    if (MailApp.getRemainingDailyQuota() <= RESET_MAIL_RESERVE + to.length) return generic;
    const code = newResetCode_();
    props.setProperties({
      ADMIN_RESET_HASH: hash_(code),
      ADMIN_RESET_EXPIRES: String(Date.now() + RESET_CODE_TTL_MS)
    });
    try {
      MailApp.sendEmail({
        to: to.join(','),
        // Not in the subject: subjects show up on lock screens, in
        // notification previews, and in forwarded threads.
        subject: 'RCCC Events — admin PIN reset',
        body: 'Someone asked to reset the admin PIN for RCCC Community Events.\n\n' +
          'Reset code: ' + code.slice(0, 4) + '-' + code.slice(4) + '\n\n' +
          'It expires in 15 minutes and works once. Enter it in the admin console.\n\n' +
          'If this was not you, ignore this email. The current PIN still works, ' +
          'nothing has changed, and whoever asked did not receive this code.'
      });
    } catch (err) {
      // Nobody got it, so nothing should be left standing waiting for it.
      clearAdminReset_(props);
      console.error('Reset code email failed: ' + err);
      return generic;
    }
    props.setProperties({ ADMIN_RESET_DAY: today, ADMIN_RESET_COUNT: String(sentToday + 1) });
    return generic;
  } finally {
    lock.releaseLock();
  }
}

function adminResetPinWithCode(code, newPin) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const expected = props.getProperty('ADMIN_RESET_HASH');
    const expires = Number(props.getProperty('ADMIN_RESET_EXPIRES') || 0);
    if (!expected || Date.now() > expires) {
      throw new Error('That code has expired, or none was issued. Ask for a new one.');
    }
    const cache = CacheService.getScriptCache();
    if (cache.get('pinreset:wait')) {
      throw new Error('Too many attempts just now. Wait a minute, then try again — your code is still good.');
    }
    if (hash_(normalizeResetCode_(code)) !== expected) {
      // Slow guessing down; never void the code. A stranger guessing must not
      // be able to take the organizer's own code away from them.
      const misses = Number(cache.get('pinreset:misses') || 0) + 1;
      if (misses >= 5) {
        cache.put('pinreset:wait', '1', 60);
        cache.remove('pinreset:misses');
      } else {
        cache.put('pinreset:misses', String(misses), 300);
      }
      throw new Error('That code is not right.');
    }
    // Validated before the code is spent, so a rejected PIN does not cost
    // another trip through the mailbox.
    applyNewAdminPin_(newPin);
    clearAdminReset_(props);
    cache.remove('pinreset:misses');
    cache.remove('pinreset:wait');
    return { ok: true, adminPinChanged: true };
  } finally {
    lock.releaseLock();
  }
}

function adminAddStation(pin, eventId, name) {
  requireAdmin_(pin);
  const event = eventById_(eventId);
  if (!event) throw new Error('Unknown event.');
  const clean = cleanText_(name);
  if (!clean) throw new Error('Enter a station name.');
  const sheet = ensureStations_(event);
  const stationId = 'S-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
  sheet.appendRow([stationId, clean, 'checkin', genStationPin_(allStationPins_()), true, '', '']);
  return adminListEvents(pin);
}

function adminCleanupLegacy(pin) {
  requireAdmin_(pin);
  cleanupLegacyData();
  return adminListEvents(pin);
}
