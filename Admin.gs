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
      events.push({
        eventId: eventId,
        name: cleanText_(row[idx.event_name]),
        date: displayDate_(row[idx.event_date]),
        location: cleanText_(row[idx.location]),
        status: cleanText_(row[idx.status]).toUpperCase(),
        registrationUrl: cleanText_(row[idx.registration_url]),
        mealScannerUrl: scannerUrl_(eventId, APP.modes.MEAL),
        spreadsheetUrl: cleanText_(row[idx.spreadsheet_url]),
        formEditUrl: cleanText_(row[idx.form_edit_url]),
        meetings: eventMeetings_(lite).map(function(m) {
          return { id: m.id, name: m.name, url: scannerUrl_(eventId, APP.modes.CHECKIN, m.id) };
        })
      });
    }
  }
  return { masterUrl: db_().getUrl(), events: events };
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
  const event = eventById_(eventId);
  if (!event) throw new Error('Unknown event.');
  if (wanted === 'finalize') {
    finalizeMealOrderFor_(event);
  } else if (wanted === 'close' || wanted === 'reopen') {
    FormApp.openById(event.formId).setAcceptingResponses(wanted === 'reopen');
    const sheet = db_().getSheetByName(SHEETS.EVENTS);
    const values = sheet.getDataRange().getValues();
    const idx = headerIndex_(values[0]);
    for (let r = 1; r < values.length; r++) {
      if (cleanText_(values[r][idx.event_id]) === event.eventId) {
        sheet.getRange(r + 1, idx.status + 1).setValue(wanted === 'reopen' ? 'ACTIVE' : 'CLOSED');
        break;
      }
    }
  } else {
    throw new Error('Unknown action "' + wanted + '".');
  }
  return adminListEvents(pin);
}

function adminAddMeeting(pin, eventId, name) {
  requireAdmin_(pin);
  const event = eventById_(eventId);
  if (!event) throw new Error('Unknown event.');
  const clean = cleanText_(name);
  if (!clean) throw new Error('Enter a meeting name.');
  const sheet = ensureSheet_(eventDb_(event), SHEETS.MEETINGS, MEETING_HEADERS, [['MAIN', 'Main check-in', true]]);
  const meetingId = 'M-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
  sheet.appendRow([meetingId, clean, true]);
  return adminListEvents(pin);
}

function adminCleanupLegacy(pin) {
  requireAdmin_(pin);
  cleanupLegacyData();
  return adminListEvents(pin);
}
