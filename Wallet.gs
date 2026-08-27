// Wallet passes. Google: a "Save to Google Wallet" link is a fat JWT signed
// RS256 with the issuer's service account — no API calls needed, the object is
// created when the guest saves it. Apple: passes are signed by the Cloudflare
// Worker at APPLE_PASS_URL; the email/receipt only link to it.
const GW_ISSUER_ID = '3388000000023178444';
const GW_CLASS_SUFFIX = 'rccc-events';

function googleWalletSaveUrl_(record, event) {
  const props = PropertiesService.getScriptProperties();
  const saEmail = props.getProperty('GW_SA_EMAIL');
  const saKey = props.getProperty('GW_SA_KEY');
  if (!saEmail || !saKey) return '';
  const objectId = GW_ISSUER_ID + '.' + (event.eventId + '-' + record.registrationId).replace(/[^\w.-]/g, '_');
  const claims = {
    iss: saEmail,
    aud: 'google',
    typ: 'savetowallet',
    iat: Math.floor(Date.now() / 1000),
    origins: [],
    payload: {
      eventTicketObjects: [{
        id: objectId,
        classId: GW_ISSUER_ID + '.' + GW_CLASS_SUFFIX,
        state: 'ACTIVE',
        ticketHolderName: record.fullName,
        ticketNumber: record.registrationId,
        hexBackgroundColor: '#1d4a38',
        barcode: { type: 'QR_CODE', value: APP.qrPrefix + record.token, alternateText: record.registrationId },
        textModulesData: [
          { header: 'Event', body: event.name },
          { header: 'Date', body: cleanText_(event.date) || '—' },
          { header: 'Location', body: cleanText_(event.location) || '—' }
        ]
      }]
    }
  };
  const input = b64url_(JSON.stringify({ alg: 'RS256', typ: 'JWT' })) + '.' + b64url_(JSON.stringify(claims));
  const signature = Utilities.computeRsaSha256Signature(input, saKey);
  return 'https://pay.google.com/gp/v/save/' + input + '.' + Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');
}

function appleWalletUrl_(record) {
  const base = PropertiesService.getScriptProperties().getProperty('APPLE_PASS_URL');
  if (!base) return '';
  return base + (base.indexOf('?') === -1 ? '?' : '&') + 'token=' + encodeURIComponent(record.token);
}

function b64url_(value) {
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(value).getBytes()).replace(/=+$/, '');
}

// Inline-styled buttons shared by the confirmation email and the receipt page.
function walletButtons_(record, event) {
  const apple = appleWalletUrl_(record);
  const google = googleWalletSaveUrl_(record, event);
  if (!apple && !google) return '';
  return '<div style="margin:4px 0 16px;text-align:center">' +
    (apple ? '<a href="' + apple + '" style="display:inline-block;margin:4px 5px;background:#000000;color:#ffffff;text-decoration:none;font-weight:bold;font-size:13px;padding:11px 18px;border-radius:10px">Add to Apple Wallet</a>' : '') +
    (google ? '<a href="' + google + '" style="display:inline-block;margin:4px 5px;background:#1a73e8;color:#ffffff;text-decoration:none;font-weight:bold;font-size:13px;padding:11px 18px;border-radius:10px">Save to Google Wallet</a>' : '') +
    '</div>';
}

// Admin-only: install the Google service-account key (paste/POST the whole
// JSON key file) and the Apple pass Worker URL.
function adminSetWalletConfig(pin, saJson, applePassUrl) {
  requireAdmin_(pin);
  const props = PropertiesService.getScriptProperties();
  if (cleanText_(saJson)) {
    const parsed = JSON.parse(saJson);
    if (!parsed.client_email || !parsed.private_key) throw new Error('That does not look like a service-account JSON key.');
    props.setProperty('GW_SA_EMAIL', parsed.client_email);
    props.setProperty('GW_SA_KEY', parsed.private_key);
  }
  if (cleanText_(applePassUrl)) {
    if (!/^https:\/\//.test(cleanText_(applePassUrl))) throw new Error('The Apple pass URL must be https.');
    props.setProperty('APPLE_PASS_URL', cleanText_(applePassUrl));
  }
  return {
    ok: true,
    googleWallet: Boolean(props.getProperty('GW_SA_EMAIL')),
    applePassUrl: props.getProperty('APPLE_PASS_URL') || ''
  };
}
