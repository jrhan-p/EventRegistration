function handleFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const named = e.namedValues || {};
    const get = function(title) { return cleanText_((named[title] || [''])[0]); };
    const now = new Date();
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    const regId = 'R-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + token.slice(0, 6).toUpperCase();
    const record = {
      createdAt: now, registrationId: regId, token: token, eventId: setting_('EVENT_ID'),
      fullName: get('Full Name'), email: get('Email Address'), phone: normalizePhone_(get('Mobile Phone')),
      mealSelected: get('Food Selection'), dietaryNotes: get('Dietary Notes'),
      smsConsent: Boolean(get('SMS Consent')), status: APP.statuses.ACTIVE
    };
    db_().getSheetByName(SHEETS.REGISTRATIONS).appendRow([
      record.createdAt, record.registrationId, record.token, record.eventId, record.fullName, record.email,
      record.phone, record.mealSelected, record.dietaryNotes, record.smsConsent, record.status,
      false, '', '', now
    ]);
    sendConfirmation_(record);
    refreshFoodSummary();
  } finally {
    lock.releaseLock();
  }
}

function sendConfirmation_(record) {
  const webAppUrl = setting_('WEB_APP_URL');
  if (!webAppUrl) throw new Error('WEB_APP_URL is not configured.');
  const receiptUrl = webAppUrl + '?view=receipt&token=' + encodeURIComponent(record.token);
  const qrValue = APP.qrPrefix + record.token;
  const qrUrl = 'https://quickchart.io/qr?size=260&text=' + encodeURIComponent(qrValue);
  const eventName = setting_('EVENT_NAME');
  const subject = 'Registration confirmed: ' + eventName;
  const html = '<div style="font-family:Arial,sans-serif;max-width:560px">' +
    '<h2>Registration confirmed</h2><p>Hello ' + html_(record.fullName) + ',</p>' +
    '<p>You are registered for <b>' + html_(eventName) + '</b>.</p>' +
    '<p><b>Date:</b> ' + html_(setting_('EVENT_DATE')) + '<br><b>Location:</b> ' + html_(setting_('LOCATION')) +
    '<br><b>Food:</b> ' + html_(record.mealSelected) + '<br><b>Registration:</b> ' + html_(record.registrationId) + '</p>' +
    '<p><img alt="Registration QR code" width="260" height="260" src="' + qrUrl + '"></p>' +
    '<p><a href="' + receiptUrl + '">Open your mobile QR receipt</a></p>' +
    '<p>Please present this QR code at check-in and meal pickup.</p></div>';
  MailApp.sendEmail({ to: record.email, subject: subject, htmlBody: html,
    body: 'Registration confirmed for ' + eventName + '. Open your QR receipt: ' + receiptUrl });
  if (record.smsConsent && record.phone) sendSms_(record.phone,
    eventName + ' registration confirmed (' + record.registrationId + '). Your QR receipt: ' + receiptUrl);
}

function sendSms_(to, body) {
  const p = PropertiesService.getScriptProperties();
  const sid = p.getProperty('TWILIO_ACCOUNT_SID');
  const token = p.getProperty('TWILIO_AUTH_TOKEN');
  const from = p.getProperty('TWILIO_FROM_NUMBER');
  if (!sid || !token || !from) {
    console.warn('SMS skipped: Twilio credentials are not configured.');
    return { skipped: true };
  }
  const response = UrlFetchApp.fetch('https://api.twilio.com/2010-04-01/Accounts/' + encodeURIComponent(sid) + '/Messages.json', {
    method: 'post', payload: { To: to, From: from, Body: body },
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token) }, muteHttpExceptions: true
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error('Twilio SMS failed: HTTP ' + response.getResponseCode() + ' ' + response.getContentText());
  }
  return { skipped: false };
}

function html_(value) {
  return cleanText_(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
