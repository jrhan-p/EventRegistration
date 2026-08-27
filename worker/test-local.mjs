// Pre-flight: build and sign a .pkpass locally with the real certificates,
// exercising the same forge/fflate code paths the Worker uses.
// Usage: node test-local.mjs
import { readFileSync, writeFileSync } from 'node:fs';
import forge from 'node-forge';
import { zipSync } from 'fflate';
import { ICONS } from './src/icons.js';

const SECRETS = '/Users/jiruihan/Documents/RCCC/wallet-secrets';
const certPem = readFileSync(`${SECRETS}/pass-cert.pem`, 'utf8');
const keyPem = readFileSync(`${SECRETS}/pass.key`, 'utf8');
const wwdrPem = readFileSync(`${SECRETS}/wwdr.pem`, 'utf8');

const data = {
  ok: true, fullName: 'Test Guest', registrationId: 'R-20260827-TEST01',
  eventId: 'EV-46497A', eventName: 'Test Lunch (safe to delete)',
  eventDate: '2026-09-01 12:00 PM', location: 'Fellowship Hall',
  qr: 'CER1:localtesttoken'
};

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
  barcodes: [{ format: 'PKBarcodeFormatQR', message: data.qr, messageEncoding: 'iso-8859-1', altText: data.registrationId }],
  eventTicket: {
    primaryFields: [{ key: 'event', label: 'EVENT', value: data.eventName }],
    secondaryFields: [
      { key: 'guest', label: 'GUEST', value: data.fullName },
      { key: 'date', label: 'DATE', value: data.eventDate }
    ],
    auxiliaryFields: [{ key: 'loc', label: 'LOCATION', value: data.location }],
    backFields: [{ key: 'reg', label: 'Registration', value: data.registrationId }]
  }
};

const files = { 'pass.json': new TextEncoder().encode(JSON.stringify(pass)) };
for (const [name, b64] of Object.entries(ICONS)) files[name] = binaryStringToBytes(atob(b64));

const manifest = {};
for (const [name, bytes] of Object.entries(files)) manifest[name] = sha1Hex(bytes);
const manifestBytes = new TextEncoder().encode(JSON.stringify(manifest));
files['manifest.json'] = manifestBytes;

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
files['signature'] = binaryStringToBytes(forge.asn1.toDer(p7.toAsn1()).getBytes());

writeFileSync('test.pkpass', zipSync(files, { level: 6 }));
writeFileSync('test-manifest.json', manifestBytes);
writeFileSync('test-signature.der', files['signature']);
console.log('WROTE test.pkpass', Object.keys(files).join(', '));

function sha1Hex(bytes) { const md = forge.md.sha1.create(); md.update(bytesToBinaryString(bytes)); return md.digest().toHex(); }
function bytesToBinaryString(bytes) { let s = ''; for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]); return s; }
function binaryStringToBytes(str) { const out = new Uint8Array(str.length); for (let i = 0; i < str.length; i++) out[i] = str.charCodeAt(i); return out; }
