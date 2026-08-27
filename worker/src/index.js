// RCCC Apple Wallet pass signer.
// GET /?token=<qr_token> -> looks the registration up through the Apps Script
// API, builds the .pkpass bundle, and signs it with the Pass Type ID
// certificate (secrets: PASS_KEY, PASS_CERT, WWDR; var: APPS_URL).
import forge from 'node-forge';
import { zipSync } from 'fflate';
import { ICONS } from './icons.js';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const token = (url.searchParams.get('token') || '').trim();
    if (!token) return new Response('Missing token', { status: 400 });

    const upstream = await fetch(env.APPS_URL, {
      method: 'POST',
      body: new URLSearchParams({ fn: 'passData', token })
    });
    const data = await upstream.json().catch(() => null);
    if (!data || !data.ok) return new Response('Registration not found', { status: 404 });

    const pass = {
      formatVersion: 1,
      passTypeIdentifier: 'pass.org.rccc.events',
      teamIdentifier: '48UG39GUGA',
      organizationName: 'RCCC Community Events',
      serialNumber: `${data.eventId}-${data.registrationId}`,
      description: `RCCC event ticket for ${data.eventName}`,
      logoText: 'RCCC Events',
      foregroundColor: 'rgb(230,201,107)',
      backgroundColor: 'rgb(29,74,56)',
      labelColor: 'rgb(201,162,39)',
      barcodes: [{
        format: 'PKBarcodeFormatQR',
        message: data.qr,
        messageEncoding: 'iso-8859-1',
        altText: data.registrationId
      }],
      eventTicket: {
        primaryFields: [{ key: 'event', label: 'EVENT', value: data.eventName }],
        secondaryFields: [
          { key: 'guest', label: 'GUEST', value: data.fullName },
          { key: 'date', label: 'DATE', value: data.eventDate || '' }
        ],
        auxiliaryFields: data.location ? [{ key: 'loc', label: 'LOCATION', value: data.location }] : [],
        backFields: [{ key: 'reg', label: 'Registration', value: data.registrationId }]
      }
    };

    const files = { 'pass.json': new TextEncoder().encode(JSON.stringify(pass)) };
    for (const [name, b64] of Object.entries(ICONS)) files[name] = b64ToBytes(b64);

    const manifest = {};
    for (const [name, bytes] of Object.entries(files)) manifest[name] = sha1Hex(bytes);
    const manifestBytes = new TextEncoder().encode(JSON.stringify(manifest));
    files['manifest.json'] = manifestBytes;
    files['signature'] = signManifest(manifestBytes, env.PASS_CERT, env.PASS_KEY, env.WWDR);

    const zip = zipSync(files, { level: 6 });
    return new Response(zip, {
      headers: {
        'content-type': 'application/vnd.apple.pkpass',
        'content-disposition': 'attachment; filename="RCCC-Event.pkpass"',
        'cache-control': 'no-store'
      }
    });
  }
};

function signManifest(manifestBytes, certPem, keyPem, wwdrPem) {
  const p7 = forge.pkcs7.createSignedData();
  p7.content = forge.util.createBuffer(bytesToBinaryString(manifestBytes));
  const cert = forge.pki.certificateFromPem(certPem);
  p7.addCertificate(forge.pki.certificateFromPem(wwdrPem));
  p7.addCertificate(cert);
  p7.addSigner({
    key: forge.pki.privateKeyFromPem(keyPem),
    certificate: cert,
    digestAlgorithm: forge.pki.oids.sha256,
    authenticatedAttributes: [
      { type: forge.pki.oids.contentType, value: forge.pki.oids.data },
      { type: forge.pki.oids.messageDigest },
      { type: forge.pki.oids.signingTime, value: new Date() }
    ]
  });
  p7.sign({ detached: true });
  return binaryStringToBytes(forge.asn1.toDer(p7.toAsn1()).getBytes());
}

function sha1Hex(bytes) {
  const md = forge.md.sha1.create();
  md.update(bytesToBinaryString(bytes));
  return md.digest().toHex();
}

function bytesToBinaryString(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}

function binaryStringToBytes(str) {
  const out = new Uint8Array(str.length);
  for (let i = 0; i < str.length; i++) out[i] = str.charCodeAt(i);
  return out;
}

function b64ToBytes(b64) {
  return binaryStringToBytes(atob(b64));
}
