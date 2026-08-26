function doGet(e) {
  const view = cleanText_(e.parameter.view || 'scanner');
  if (view === 'receipt') return renderReceipt_(e.parameter.token);
  const template = HtmlService.createTemplateFromFile('ScannerPage');
  template.mode = e.parameter.mode === APP.modes.MEAL ? APP.modes.MEAL : APP.modes.CHECKIN;
  template.eventName = setting_('EVENT_NAME');
  return template.evaluate().setTitle('Event scanner').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderReceipt_(token) {
  const found = findRegistration_(cleanText_(token));
  if (!found) return HtmlService.createHtmlOutput('<h2>Registration not found</h2>');
  const r = rowToRecord_(found.values, found.index);
  const template = HtmlService.createTemplateFromFile('Receipt');
  template.record = r;
  template.eventName = setting_('EVENT_NAME');
  template.eventDate = setting_('EVENT_DATE');
  template.location = setting_('LOCATION');
  template.qrUrl = 'https://quickchart.io/qr?size=320&text=' + encodeURIComponent(APP.qrPrefix + cleanText_(token));
  return template.evaluate().setTitle('Registration receipt').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
