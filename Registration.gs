function handleFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const event = eventBySpreadsheetId_(e.range.getSheet().getParent().getId());
    if (!event) {
      console.error('Form submission arrived from a spreadsheet that matches no event; ignoring.');
      return;
    }
    const named = e.namedValues || {};
    const get = function(title) { return cleanText_((named[title] || [''])[0]); };
    const now = new Date();
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    const regId = 'R-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + token.slice(0, 6).toUpperCase();
    const record = {
      createdAt: now, registrationId: regId, token: token, eventId: event.eventId,
      fullName: get('Full Name'), email: firstEmailValue_(named), phone: normalizePhone_(get('Mobile Phone')),
      mealSelected: get('Food Selection'), dietaryNotes: get('Dietary Notes'),
      smsConsent: Boolean(get('SMS Consent')), status: APP.statuses.ACTIVE
    };
    record.userId = findOrCreateUser_(record.email, record.fullName, record.phone, now);
    // Self-service updates: a resubmission by the same signed-in user for the
    // same attendee name overwrites the earlier registration in place — the
    // guest keeps the same QR code, and a changed meal is no longer "ordered".
    const existing = findExistingRegistration_(event, record.userId, record.fullName);
    if (existing) {
      const i = existing.index;
      const prevMeal = cleanText_(existing.values[i.meal_selected]);
      // phone / meal_selected / dietary_notes / sms_consent are adjacent
      // columns in REG_HEADERS, written in one call.
      existing.sheet.getRange(existing.rowNumber, i.phone + 1, 1, 4)
        .setValues([[record.phone, record.mealSelected, record.dietaryNotes, record.smsConsent]]);
      if (prevMeal !== record.mealSelected && bool_(existing.values[i.meal_ordered])) {
        existing.sheet.getRange(existing.rowNumber, i.meal_ordered + 1).setValue(false);
      }
      existing.sheet.getRange(existing.rowNumber, i.updated_at + 1).setValue(now);
      record.registrationId = cleanText_(existing.values[i.registration_id]);
      record.token = cleanText_(existing.values[i.qr_token]);
      record.updated = true;
      try {
        sendConfirmation_(record, event);
      } catch (err) {
        console.error('Update confirmation delivery failed for ' + record.registrationId + ': ' + err);
      }
      refreshFoodSummaryFor_(event);
      return;
    }
    eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS).appendRow([
      record.createdAt, record.registrationId, record.token, record.eventId, record.fullName, record.email,
      record.phone, record.mealSelected, record.dietaryNotes, record.smsConsent, record.status,
      false, '', '', now, record.userId
    ]);
    // A delivery failure must never lose the registration itself.
    try {
      sendConfirmation_(record, event);
    } catch (err) {
      console.error('Confirmation delivery failed for ' + record.registrationId + ': ' + err);
    }
    refreshFoodSummaryFor_(event);
  } finally {
    lock.releaseLock();
  }
}

// The Google-verified respondent email. Header casing varies by Forms version,
// so scan every email-ish column for the first non-empty value.
function firstEmailValue_(named) {
  const keys = Object.keys(named || {});
  for (let i = 0; i < keys.length; i++) {
    if (!/^email/i.test(cleanText_(keys[i]))) continue;
    const values = named[keys[i]] || [];
    for (let j = 0; j < values.length; j++) {
      const value = cleanText_(values[j]);
      if (value) return value;
    }
  }
  return '';
}

// Editor maintenance helper: resend the confirmation for the newest active
// registration across all events.
function resendLatestConfirmation() {
  let newest = null;
  allEvents_().forEach(function(event) {
    const values = eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS).getDataRange().getValues();
    if (values.length < 2) return;
    const idx = headerIndex_(values[0]);
    for (let r = values.length - 1; r >= 1; r--) {
      if (values[r][idx.status] !== APP.statuses.ACTIVE) continue;
      const createdAt = new Date(values[r][idx.created_at]);
      if (newest && createdAt <= newest.createdAt) break;
      newest = {
        createdAt: createdAt,
        event: event,
        record: {
          registrationId: cleanText_(values[r][idx.registration_id]),
          token: cleanText_(values[r][idx.qr_token]),
          fullName: cleanText_(values[r][idx.full_name]),
          email: cleanText_(values[r][idx.email]),
          phone: cleanText_(values[r][idx.phone]),
          mealSelected: cleanText_(values[r][idx.meal_selected]),
          smsConsent: bool_(values[r][idx.sms_consent])
        }
      };
      break;
    }
  });
  if (!newest) throw new Error('No active registrations exist.');
  if (!newest.record.email) throw new Error('Registration ' + newest.record.registrationId + ' has no email recorded; submit the form again instead.');
  sendConfirmation_(newest.record, newest.event);
  console.log('Confirmation resent for ' + newest.record.registrationId + ' to ' + newest.record.email);
}

function findExistingRegistration_(event, userId, fullName) {
  const sheet = eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getDataRange().getValues();
  const idx = headerIndex_(values[0]);
  const wanted = cleanText_(fullName).toLowerCase();
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r][idx.status] !== APP.statuses.ACTIVE) continue;
    if (cleanText_(values[r][idx.user_id]) === userId &&
        cleanText_(values[r][idx.full_name]).toLowerCase() === wanted) {
      return { sheet: sheet, values: values[r], index: idx, rowNumber: r + 1 };
    }
  }
  return null;
}

// Strictly deterministic: one verified Google email maps to exactly one user_id.
// No matching by phone or name — ambiguous cases stay separate users on purpose.
function findOrCreateUser_(email, name, phone, now) {
  const sheet = ensureSheet_(db_(), SHEETS.USERS, USER_HEADERS);
  const values = sheet.getDataRange().getValues();
  const idx = headerIndex_(values[0] && values[0].length > 1 ? values[0] : USER_HEADERS.slice());
  const wanted = cleanText_(email).toLowerCase();
  for (let r = 1; r < values.length; r++) {
    if (cleanText_(values[r][idx.google_email]).toLowerCase() === wanted && wanted) {
      sheet.getRange(r + 1, idx.updated_at + 1).setValue(now);
      return cleanText_(values[r][idx.user_id]);
    }
  }
  const userId = 'U-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  sheet.appendRow([userId, email, name, phone, now, now]);
  return userId;
}

function sendConfirmation_(record, event) {
  const webAppUrl = setting_('WEB_APP_URL');
  if (!webAppUrl) throw new Error('WEB_APP_URL is not configured.');
  const receiptUrl = webAppUrl + '?view=receipt&event=' + encodeURIComponent(event.eventId) + '&token=' + encodeURIComponent(record.token);
  const qrValue = APP.qrPrefix + record.token;
  const qrUrl = 'https://quickchart.io/qr?size=260&text=' + encodeURIComponent(qrValue);
  const subject = (record.updated ? 'Registration updated: ' : 'Registration confirmed: ') + event.name;
  const updatedNote = record.updated
    ? '<tr><td style="font-size:12px;color:#1a7f4e;padding:0 22px 10px;font-family:Arial,Helvetica,sans-serif" align="center">Your registration was updated &mdash; the same QR code still works.</td></tr>'
    : '';
  const label = 'font-size:10px;letter-spacing:2px;text-transform:uppercase;color:#8a8272';
  // The ticket deliberately shows no entitlements (meal, sessions, …):
  // stations resolve those at scan time. The pass is identity + credential only.
  const html =
    '<div style="background:#faf6ec;padding:18px 0">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="width:400px;max-width:94%;background:#ffffff;border-radius:16px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;color:#26221a">' +
    '<tr><td style="background:#1d4a38;padding:16px 22px 13px">' +
    '<div style="color:#e6c96b;font-size:10px;letter-spacing:3px;text-transform:uppercase">&#10022;&nbsp; RCCC Community Events</div>' +
    '<div style="color:#ffffff;font-family:Georgia,\'Times New Roman\',serif;font-size:20px;padding-top:6px">' + html_(event.name) + '</div>' +
    '<div style="color:#bcd2c5;font-size:12px;padding-top:3px">' + html_(event.date) + (event.location ? ' &middot; ' + html_(event.location) : '') + '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:14px 22px 14px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td width="55%" align="left" style="vertical-align:top">' +
    '<div style="' + label + '">Guest</div>' +
    '<div style="font-family:Georgia,serif;font-size:19px;padding-top:2px">' + html_(record.fullName) + '</div></td>' +
    '<td width="45%" align="right" style="vertical-align:top">' +
    '<div style="' + label + '">Registration</div>' +
    '<div style="font-size:12px;font-family:Courier,monospace;padding-top:5px;white-space:nowrap">' + html_(record.registrationId) + '</div></td>' +
    '</tr></table>' +
    '</td></tr>' +
    updatedNote +
    '<tr><td align="center" style="border-top:2px dashed #ddd6c4;padding:14px 22px 16px">' +
    '<img src="' + qrUrl + '" width="180" height="180" alt="Registration QR code" style="display:block;margin:0 auto;border:0">' +
    '<div style="font-size:12px;color:#8a8272;padding-top:8px">Present this code at check-in and meal pickup.</div>' +
    walletButtons_(record, event) +
    '<div style="padding-top:2px"><a href="' + receiptUrl + '" style="color:#8a5a00;font-size:12px">View this pass in your browser</a></div>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table></div>';
  MailApp.sendEmail({ to: record.email, subject: subject, htmlBody: html, name: 'RCCC Events',
    body: 'Registration confirmed for ' + event.name + '. Open your QR receipt: ' + receiptUrl });
  if (record.smsConsent && record.phone) sendSms_(record.phone,
    event.name + ' registration confirmed (' + record.registrationId + '). Your QR receipt: ' + receiptUrl);
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
