const SHEETS = Object.freeze({
  SETTINGS: 'Settings',
  MEALS: 'Meals',
  REGISTRATIONS: 'Registrations',
  SCAN_LOG: 'Scan Log',
  FOOD_SUMMARY: 'Food Summary'
});

const REG_HEADERS = Object.freeze([
  'created_at', 'registration_id', 'qr_token', 'event_id', 'full_name', 'email',
  'phone', 'meal_selected', 'dietary_notes', 'sms_consent', 'status',
  'meal_ordered', 'checked_in_at', 'meal_redeemed_at', 'updated_at'
]);

const LOG_HEADERS = Object.freeze([
  'timestamp', 'event_id', 'registration_id', 'mode', 'result_code', 'operator', 'details'
]);

function setupApplication() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.create('Church Event Registration');
  spreadsheetId = ss.getId();
  props.setProperty('SPREADSHEET_ID', spreadsheetId);

  ensureSheet_(ss, SHEETS.SETTINGS, ['key', 'value'], [
    ['EVENT_ID', 'event-001'], ['EVENT_NAME', 'Church Community Event'],
    ['EVENT_DATE', '2026-10-01 10:00 AM'], ['LOCATION', 'Main Hall'],
    ['ORGANIZER_EMAIL', Session.getEffectiveUser().getEmail()], ['WEB_APP_URL', '']
  ]);
  ensureSheet_(ss, SHEETS.MEALS, ['meal_name', 'active'], [
    ['Chicken', true], ['Vegetarian', true], [APP.noMeal, true]
  ]);
  ensureSheet_(ss, SHEETS.REGISTRATIONS, REG_HEADERS);
  ensureSheet_(ss, SHEETS.SCAN_LOG, LOG_HEADERS);
  ensureSheet_(ss, SHEETS.FOOD_SUMMARY, ['meal_name', 'requested', 'marked_ordered', 'redeemed']);

  let formId = props.getProperty('FORM_ID');
  const form = formId ? FormApp.openById(formId) : FormApp.create('Church Event Registration');
  props.setProperty('FORM_ID', form.getId());
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  form.setAcceptingResponses(false);
  buildFormItems_(form);
  installSubmitTrigger_(ss);
  refreshFoodSummary();
  return { spreadsheetUrl: ss.getUrl(), formEditUrl: form.getEditUrl(), formPublicUrl: form.getPublishedUrl() };
}

function setOperatorPins(checkInPin, mealPin) {
  if (!/^\d{4,10}$/.test(String(checkInPin)) || !/^\d{4,10}$/.test(String(mealPin))) {
    throw new Error('Each operator PIN must contain 4 to 10 digits.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CHECKIN_PIN_HASH', hash_(String(checkInPin)));
  props.setProperty('MEAL_PIN_HASH', hash_(String(mealPin)));
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

function activateRegistration() {
  const url = setting_('WEB_APP_URL');
  if (!url) throw new Error('Deploy the web app and call setWebAppUrl(url) first.');
  if (!PropertiesService.getScriptProperties().getProperty('CHECKIN_PIN_HASH')) throw new Error('Call setOperatorPins first.');
  syncMealChoicesToForm();
  FormApp.openById(PropertiesService.getScriptProperties().getProperty('FORM_ID')).setAcceptingResponses(true);
}

function syncMealChoicesToForm() {
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty('FORM_ID'));
  const choices = activeMeals_();
  if (!choices.length) throw new Error('Add at least one active meal in the Meals sheet.');
  const item = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE).map(function(i) { return i.asMultipleChoiceItem(); })
    .filter(function(i) { return i.getTitle() === 'Food Selection'; })[0];
  if (!item) throw new Error('Food Selection field was not found.');
  item.setChoiceValues(choices);
}

function finalizeMealOrder() {
  const sheet = db_().getSheetByName(SHEETS.REGISTRATIONS);
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
  refreshFoodSummary();
}

function buildFormItems_(form) {
  if (form.getItems().length) return;
  form.setDescription('One registration and one QR code per attendee.');
  form.addTextItem().setTitle('Full Name').setRequired(true);
  form.addTextItem().setTitle('Email Address').setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().build());
  form.addTextItem().setTitle('Mobile Phone').setHelpText('Include country code when outside the U.S.').setRequired(true);
  form.addMultipleChoiceItem().setTitle('Food Selection').setChoiceValues(activeMeals_()).setRequired(true);
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

function installSubmitTrigger_(ss) {
  ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'handleFormSubmit'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('handleFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
}

function activeMeals_() {
  const values = db_().getSheetByName(SHEETS.MEALS).getDataRange().getValues();
  return values.slice(1).filter(function(r) { return bool_(r[1]); }).map(function(r) { return cleanText_(r[0]); }).filter(Boolean);
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
