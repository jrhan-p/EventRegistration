const SHEETS = Object.freeze({
  SETTINGS: 'Settings',
  USERS: 'Users',
  EVENTS: 'Events',
  MEAL_TEMPLATE: 'Meal Template',
  MEALS: 'Meals',
  REGISTRATIONS: 'Registrations',
  SCAN_LOG: 'Scan Log',
  FOOD_SUMMARY: 'Food Summary',
  STATIONS: 'Stations',
  CHECKINS: 'Check-ins'
});

// Every operator post is a "station": the MAIN check-in desk, the meal-pickup
// window, and any extra meeting/session check-ins. Each station has its own
// PIN, and the PIN alone routes the universal scanner page to the right
// station — closing an event silently disables all of its PINs. PINs are
// stored in the event's own spreadsheet so organizers can read and change them.
const STATION_HEADERS = Object.freeze(['station_id', 'station_name', 'type', 'pin', 'active']);
const CHECKIN_HEADERS = Object.freeze(['timestamp', 'station_id', 'registration_id', 'operator']);

const REG_HEADERS = Object.freeze([
  'created_at', 'registration_id', 'qr_token', 'event_id', 'full_name', 'email',
  'phone', 'meal_selected', 'dietary_notes', 'sms_consent', 'status',
  'meal_ordered', 'checked_in_at', 'meal_redeemed_at', 'updated_at', 'user_id'
]);

const LOG_HEADERS = Object.freeze([
  'timestamp', 'event_id', 'registration_id', 'mode', 'result_code', 'operator', 'details'
]);

const USER_HEADERS = Object.freeze([
  'user_id', 'google_email', 'display_name', 'phone', 'created_at', 'updated_at'
]);

// The Events sheet is the organizer's control surface: add a row with an
// event_name and run applyEventChanges() to provision it; type finalize/close/
// reopen in the action column and run applyEventChanges() to operate it.
// The camera scanner is hosted on GitHub Pages: Google's Apps Script iframe
// wrapper denies camera access (NotAllowedError on iOS WebKit), so the scanner
// must be served top-level. It talks back to this script through the doPost API.
const SCANNER_BASE = 'https://jrhan-p.github.io/EventRegistration/scanner.html';

function eventStations_(event) {
  try {
    const sheet = eventDb_(event).getSheetByName(SHEETS.STATIONS);
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    const idx = headerIndex_(values[0]);
    return values.slice(1)
      .filter(function(r) { return bool_(r[idx.active]) && cleanText_(r[idx.station_id]); })
      .map(function(r) {
        return {
          id: cleanText_(r[idx.station_id]),
          name: cleanText_(r[idx.station_name]),
          type: cleanText_(r[idx.type]).toLowerCase() === 'meal' ? 'meal' : 'checkin',
          pin: cleanText_(r[idx.pin])
        };
      });
  } catch (err) {
    return [];
  }
}

function ensureStations_(event) {
  const ss = eventDb_(event);
  let sheet = ss.getSheetByName(SHEETS.STATIONS);
  if (sheet) return sheet;
  const taken = allStationPins_();
  return ensureSheet_(ss, SHEETS.STATIONS, STATION_HEADERS, [
    ['MAIN', 'Main check-in', 'checkin', genStationPin_(taken), true],
    ['MEAL', 'Meal pickup', 'meal', genStationPin_(taken), true]
  ]);
}

// The PIN alone identifies the station. Only ACTIVE events participate, so
// closing an event retires its PINs. The full lookup opens every active
// event's spreadsheet, so resolved stations are cached for five minutes —
// scans hit the cache and skip all of that. (A closed event's PINs can
// therefore linger for up to five minutes.)
function resolvePin_(pin) {
  const clean = cleanText_(pin);
  if (!clean) return null;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('pin:' + clean);
  if (cached) {
    try {
      const station = JSON.parse(cached);
      if (station && station.event && station.event.spreadsheetId) return station;
    } catch (err) {}
  }
  const events = allEvents_();
  for (let i = 0; i < events.length; i++) {
    if (events[i].status !== 'ACTIVE') continue;
    const stations = eventStations_(events[i]);
    for (let s = 0; s < stations.length; s++) {
      if (stations[s].pin && stations[s].pin === clean) {
        const station = { event: events[i], stationId: stations[s].id, stationName: stations[s].name, type: stations[s].type };
        cache.put('pin:' + clean, JSON.stringify(station), 300);
        return station;
      }
    }
  }
  return null;
}

function allStationPins_() {
  const taken = {};
  allEvents_().forEach(function(event) {
    eventStations_(event).forEach(function(s) { if (s.pin) taken[s.pin] = true; });
  });
  return taken;
}

function genStationPin_(taken) {
  let pin;
  do { pin = String(Math.floor(Math.random() * 900000) + 100000); } while (taken[pin]);
  taken[pin] = true;
  return pin;
}

const EVENT_HEADERS = Object.freeze([
  'event_id', 'event_name', 'event_date', 'location', 'status', 'action',
  'registration_url', 'checkin_scanner_url', 'meal_scanner_url',
  'spreadsheet_url', 'form_edit_url', 'spreadsheet_id', 'form_id', 'created_at'
]);

function setupApplication() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.create('Church Event Registration – Master');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  ensureSheet_(ss, SHEETS.SETTINGS, ['key', 'value'], [
    ['ORGANIZER_EMAIL', Session.getEffectiveUser().getEmail()], ['WEB_APP_URL', '']
  ]);
  ensureSheet_(ss, SHEETS.USERS, USER_HEADERS);
  ensureSheet_(ss, SHEETS.EVENTS, EVENT_HEADERS);
  ensureSheet_(ss, SHEETS.MEAL_TEMPLATE, ['meal_name', 'active'], [
    ['Chicken', true], ['Vegetarian', true], [APP.noMeal, true]
  ]);
  return { masterSpreadsheetUrl: ss.getUrl() };
}

function applyEventChanges() {
  const sheet = db_().getSheetByName(SHEETS.EVENTS);
  if (!sheet) throw new Error('Run setupApplication() first.');
  const values = sheet.getDataRange().getValues();
  const idx = headerIndex_(values[0]);
  const messages = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!cleanText_(row[idx.event_name])) continue;
    const status = cleanText_(row[idx.status]).toUpperCase();
    const action = cleanText_(row[idx.action]).toLowerCase();
    if (!status || status === 'NEW') {
      messages.push(provisionEvent_(sheet, idx, r + 1, row));
      continue;
    }
    if (!action) continue;
    const event = eventFromRow_(row, idx);
    if (action === 'finalize') {
      finalizeMealOrderFor_(event);
      messages.push(event.eventId + ': meal order finalized');
    } else if (action === 'close') {
      FormApp.openById(event.formId).setAcceptingResponses(false);
      sheet.getRange(r + 1, idx.status + 1).setValue('CLOSED');
      messages.push(event.eventId + ': registration closed');
    } else if (action === 'reopen') {
      FormApp.openById(event.formId).setAcceptingResponses(true);
      sheet.getRange(r + 1, idx.status + 1).setValue('ACTIVE');
      messages.push(event.eventId + ': registration reopened');
    } else {
      messages.push(event.eventId + ': unknown action "' + action + '" (use finalize, close, or reopen)');
    }
    sheet.getRange(r + 1, idx.action + 1).setValue('');
  }
  if (!messages.length) {
    messages.push('Nothing to do. Add a row with an event_name, or type finalize/close/reopen in the action column.');
  }
  messages.forEach(function(m) { console.log(m); });
  return messages;
}

function provisionEvent_(eventsSheet, idx, rowNumber, row) {
  const name = cleanText_(row[idx.event_name]);
  let eventId = cleanText_(row[idx.event_id]);
  if (!eventId) {
    eventId = 'EV-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
    eventsSheet.getRange(rowNumber, idx.event_id + 1).setValue(eventId);
  }
  const ss = SpreadsheetApp.create('Event Registration – ' + name);
  const initialSheet = ss.getSheets()[0];
  const meals = templateMeals_();
  ensureSheet_(ss, SHEETS.MEALS, ['meal_name', 'active'], meals.map(function(m) { return [m, true]; }));
  ensureSheet_(ss, SHEETS.REGISTRATIONS, REG_HEADERS);
  ensureSheet_(ss, SHEETS.SCAN_LOG, LOG_HEADERS);
  ensureSheet_(ss, SHEETS.FOOD_SUMMARY, ['meal_name', 'requested', 'marked_ordered', 'redeemed']);
  const takenPins = allStationPins_();
  ensureSheet_(ss, SHEETS.STATIONS, STATION_HEADERS, [
    ['MAIN', 'Main check-in', 'checkin', genStationPin_(takenPins), true],
    ['MEAL', 'Meal pickup', 'meal', genStationPin_(takenPins), true]
  ]);
  ensureSheet_(ss, SHEETS.CHECKINS, CHECKIN_HEADERS);
  if (ss.getSheets().length > 1) ss.deleteSheet(initialSheet);
  const form = FormApp.create(name + ' – Registration');
  requireVerifiedGoogleEmail_(form);
  buildFormItems_(form, meals);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.setAcceptingResponses(true);
  ScriptApp.newTrigger('handleFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  eventsSheet.getRange(rowNumber, idx.status + 1).setValue('ACTIVE');
  eventsSheet.getRange(rowNumber, idx.registration_url + 1).setValue(form.getPublishedUrl());
  eventsSheet.getRange(rowNumber, idx.checkin_scanner_url + 1).setValue(SCANNER_BASE);
  eventsSheet.getRange(rowNumber, idx.meal_scanner_url + 1).setValue(SCANNER_BASE);
  eventsSheet.getRange(rowNumber, idx.spreadsheet_url + 1).setValue(ss.getUrl());
  eventsSheet.getRange(rowNumber, idx.form_edit_url + 1).setValue(form.getEditUrl());
  eventsSheet.getRange(rowNumber, idx.spreadsheet_id + 1).setValue(ss.getId());
  eventsSheet.getRange(rowNumber, idx.form_id + 1).setValue(form.getId());
  eventsSheet.getRange(rowNumber, idx.created_at + 1).setValue(new Date());
  return eventId + ': provisioned (' + name + ')';
}

// Removes the single-event tabs, test users, and legacy form left over from the
// pre-multi-event deployment. Safe to run more than once.
function cleanupLegacyData() {
  const ss = db_();
  const props = PropertiesService.getScriptProperties();
  const legacyFormId = props.getProperty('FORM_ID');
  if (legacyFormId) {
    try {
      const form = FormApp.openById(legacyFormId);
      form.setAcceptingResponses(false);
      try { form.removeDestination(); } catch (err) {}
      form.setTitle('[OLD – safe to delete] ' + form.getTitle().replace(/^\[OLD – safe to delete\] /, ''));
    } catch (err) {
      console.warn('Legacy form could not be updated: ' + err);
    }
    props.deleteProperty('FORM_ID');
  }
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'handleFormSubmit' && t.getTriggerSourceId() === ss.getId()) {
      ScriptApp.deleteTrigger(t);
    }
  });
  [SHEETS.REGISTRATIONS, SHEETS.MEALS, SHEETS.SCAN_LOG, SHEETS.FOOD_SUMMARY].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
  ss.getSheets().forEach(function(sheet) {
    if (/^Form Responses/i.test(sheet.getName())) ss.deleteSheet(sheet);
  });
  const users = ss.getSheetByName(SHEETS.USERS);
  if (users && users.getLastRow() > 1) users.deleteRows(2, users.getLastRow() - 1);
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (settings) {
    const rows = settings.getDataRange().getValues();
    for (let r = rows.length - 1; r >= 1; r--) {
      if (['EVENT_ID', 'EVENT_NAME', 'EVENT_DATE', 'LOCATION'].indexOf(cleanText_(rows[r][0])) !== -1) {
        settings.deleteRow(r + 1);
      }
    }
  }
  console.log('Legacy data removed. The old form is closed and renamed "[OLD – safe to delete]"; delete it from Drive whenever you like.');
}

function setTwilioCredentials(accountSid, authToken, fromNumber) {
  PropertiesService.getScriptProperties().setProperties({
    TWILIO_ACCOUNT_SID: cleanText_(accountSid),
    TWILIO_AUTH_TOKEN: cleanText_(authToken),
    TWILIO_FROM_NUMBER: normalizePhone_(fromNumber)
  });
}

function setWebAppUrl(url) {
  const clean = cleanText_(url);
  if (!/^https:\/\/script\.google\.com\//.test(clean)) throw new Error('Enter the deployed Apps Script web app URL.');
  setSetting_('WEB_APP_URL', clean);
}

function requireVerifiedGoogleEmail_(form) {
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
  } catch (err) {
    form.setCollectEmail(true);
  }
}

// Pushes each active event's Meals sheet back into its form's Food Selection
// choices. Run after editing an event's Meals sheet.
function syncAllMealChoices() {
  allEvents_().forEach(function(event) {
    if (event.status !== 'ACTIVE') return;
    const meals = activeMealsFor_(event);
    if (!meals.length) return;
    const form = FormApp.openById(event.formId);
    const item = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE)
      .map(function(i) { return i.asMultipleChoiceItem(); })
      .filter(function(i) { return i.getTitle() === 'Food Selection'; })[0];
    if (item) item.setChoiceValues(meals);
  });
}

function finalizeMealOrderFor_(event) {
  const sheet = eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const idx = headerIndex_(values[0]);
  const now = new Date();
  for (let r = 1; r < values.length; r++) {
    const active = values[r][idx.status] === APP.statuses.ACTIVE;
    const hasMeal = cleanText_(values[r][idx.meal_selected]) !== APP.noMeal;
    if (active && hasMeal) {
      values[r][idx.meal_ordered] = true;
      values[r][idx.updated_at] = now;
    }
  }
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  refreshFoodSummaryFor_(event);
}

function buildFormItems_(form, meals) {
  if (form.getItems().length) return;
  form.setDescription('One registration and one QR code per attendee. Registering for family members is fine: submit the form once per person.');
  form.addTextItem().setTitle('Full Name').setRequired(true);
  form.addTextItem().setTitle('Mobile Phone').setHelpText('Include country code when outside the U.S.').setRequired(true);
  form.addMultipleChoiceItem().setTitle('Food Selection').setChoiceValues(meals).setRequired(true);
  form.addParagraphTextItem().setTitle('Dietary Notes').setRequired(false);
  form.addCheckboxItem().setTitle('SMS Consent')
    .setHelpText('I agree to receive registration-related text messages for this event. Message and data rates may apply.')
    .setChoiceValues(['I agree']).setRequired(true);
}

function ensureSheet_(ss, name, headers, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
    sheet.setFrozenRows(1);
    if (rows && rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function templateMeals_() {
  const sheet = db_().getSheetByName(SHEETS.MEAL_TEMPLATE);
  if (!sheet) throw new Error('Run setupApplication() first.');
  const meals = sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return bool_(r[1]); })
    .map(function(r) { return cleanText_(r[0]); })
    .filter(Boolean);
  if (!meals.length) throw new Error('Add at least one active meal in the Meal Template sheet.');
  return meals;
}

function activeMealsFor_(event) {
  const sheet = eventDb_(event).getSheetByName(SHEETS.MEALS);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return bool_(r[1]); })
    .map(function(r) { return cleanText_(r[0]); })
    .filter(Boolean);
}

function allEvents_() {
  const sheet = db_().getSheetByName(SHEETS.EVENTS);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const idx = headerIndex_(values[0]);
  return values.slice(1)
    .map(function(row) { return eventFromRow_(row, idx); })
    .filter(function(event) { return event.eventId && event.spreadsheetId; });
}

function eventById_(eventId) {
  const wanted = cleanText_(eventId);
  return allEvents_().filter(function(event) { return event.eventId === wanted; })[0] || null;
}

function eventBySpreadsheetId_(spreadsheetId) {
  return allEvents_().filter(function(event) { return event.spreadsheetId === spreadsheetId; })[0] || null;
}

function eventFromRow_(row, idx) {
  return {
    eventId: cleanText_(row[idx.event_id]),
    name: cleanText_(row[idx.event_name]),
    date: displayDate_(row[idx.event_date]),
    location: cleanText_(row[idx.location]),
    status: cleanText_(row[idx.status]).toUpperCase(),
    spreadsheetId: cleanText_(row[idx.spreadsheet_id]),
    formId: cleanText_(row[idx.form_id])
  };
}

function eventDb_(event) {
  return SpreadsheetApp.openById(event.spreadsheetId);
}

function displayDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd h:mm a');
  return cleanText_(value);
}

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run setupApplication() first.');
  return SpreadsheetApp.openById(id);
}

function setting_(key) {
  const values = db_().getSheetByName(SHEETS.SETTINGS).getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (values[i][0] === key) return cleanText_(values[i][1]);
  return '';
}

function setSetting_(key, value) {
  const sheet = db_().getSheetByName(SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function hash_(value) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value));
}

function headerIndex_(headers) {
  return headers.reduce(function(out, h, i) { out[cleanText_(h)] = i; return out; }, {});
}
