// The PIN alone routes the scan: it identifies the event, the station, and
// whether this is a check-in desk or the meal window.
//
// Latency design: the registration is located OUTSIDE the lock using cached
// row positions; only the duplicate check and the cell writes are serialized,
// so several stations can scan concurrently without queueing behind each
// other's full scan.
function processScan(scanValue, pin) {
  const token = extractToken_(scanValue);
  const station = resolvePin_(pin);
  if (!station) return { ok: false, code: 'UNAUTHORIZED', message: 'Operator PIN is incorrect.' };
  const event = station.event;
  const mode = station.type === 'meal' ? APP.modes.MEAL : APP.modes.CHECKIN;
  const now = new Date();
  const found = findRegistration_(token, event);
  if (!found) {
    const miss = evaluateScan_(null, mode);
    logScan_(event, null, mode + '@' + station.stationId, miss, now);
    return miss;
  }
  const record = rowToRecord_(found.values, found.index);
  const index = found.index;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Re-read just the volatile cells inside the lock so a concurrent scan
    // cannot double-pass the same guest.
    const cols = [index.meal_ordered, index.checked_in_at, index.meal_redeemed_at, index.updated_at];
    const lo = Math.min.apply(null, cols);
    const hi = Math.max.apply(null, cols);
    const span = found.sheet.getRange(found.rowNumber, lo + 1, 1, hi - lo + 1);
    const vol = span.getValues()[0];
    record.mealOrdered = bool_(vol[index.meal_ordered - lo]);
    record.mealRedeemedAt = vol[index.meal_redeemed_at - lo] || '';
    if (mode === APP.modes.CHECKIN) {
      // Fast path: an empty checked_in_at means no station anywhere has seen
      // this guest, so the per-station duplicate lookup can be skipped.
      record.checkedInAt = vol[index.checked_in_at - lo]
        ? stationCheckin_(event, station.stationId, record.registrationId) : '';
    }
    const decision = evaluateScan_(record, mode);
    if (decision.ok) {
      if (mode === APP.modes.CHECKIN) {
        if (!vol[index.checked_in_at - lo]) vol[index.checked_in_at - lo] = now;
        checkinsSheet_(event).appendRow([now, station.stationId, record.registrationId, Session.getActiveUser().getEmail() || 'PIN operator']);
        CacheService.getScriptCache().put(
          'dup:' + event.spreadsheetId + ':' + station.stationId + ':' + record.registrationId,
          now.toISOString(), 21600);
      } else {
        vol[index.meal_redeemed_at - lo] = now;
      }
      vol[index.updated_at - lo] = now;
      span.setValues([vol]);
    }
    logScan_(event, record, mode + '@' + station.stationId, decision, now);
    if (decision.ok && mode === APP.modes.MEAL) bumpFoodSummary_(event, record.mealSelected);
    // The check-in desk must not see meal choices; only the kitchen scanner may.
    const extra = { registrationId: record.registrationId, fullName: record.fullName, timestamp: now.toISOString() };
    if (mode === APP.modes.MEAL) {
      extra.mealSelected = record.mealSelected;
      extra.dietaryNotes = record.dietaryNotes;
    }
    return Object.assign({}, decision, extra);
  } finally {
    lock.releaseLock();
  }
}

// Front-desk fallback for guests who cannot find their QR code. Requires a
// station PIN; matches on the trailing digits of the phone number.
function phoneLookup(pin, phone) {
  const station = resolvePin_(pin);
  if (!station) return { ok: false, message: 'Operator PIN is incorrect.' };
  const digits = cleanText_(phone).replace(/\D/g, '');
  if (digits.length < 4) return { ok: false, message: 'Enter at least the last 4 digits of the phone number.' };
  const sheet = eventDb_(station.event).getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, matches: [] };
  const values = sheet.getDataRange().getValues();
  const idx = headerIndex_(values[0]);
  const matches = [];
  for (let r = 1; r < values.length && matches.length < 6; r++) {
    if (values[r][idx.status] !== APP.statuses.ACTIVE) continue;
    const p = cleanText_(values[r][idx.phone]).replace(/\D/g, '');
    if (p && p.slice(-digits.length) === digits) {
      matches.push({
        fullName: cleanText_(values[r][idx.full_name]),
        registrationId: cleanText_(values[r][idx.registration_id]),
        token: cleanText_(values[r][idx.qr_token])
      });
    }
  }
  return { ok: true, matches: matches };
}

// Reads only what it must: cached header index and cached token->row position
// mean a warm lookup is a single-row read. Rows are append-only, so cached
// positions stay valid; a mismatch (manual row deletion) clears and rescans.
function findRegistration_(token, event) {
  const sheet = eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet) return null;
  const cache = CacheService.getScriptCache();
  const ssid = event.spreadsheetId;
  const lastCol = sheet.getLastColumn();
  let index = null;
  const hdrHit = cache.get('hdr:' + ssid);
  if (hdrHit) {
    try { index = JSON.parse(hdrHit); } catch (err) {}
  }
  if (!index || index.qr_token === undefined) {
    index = headerIndex_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
    cache.put('hdr:' + ssid, JSON.stringify(index), 21600);
  }
  const key = 'tok:' + ssid + ':' + token;
  const cachedRow = Number(cache.get(key)) || 0;
  if (cachedRow) {
    const values = sheet.getRange(cachedRow, 1, 1, lastCol).getValues()[0];
    if (cleanText_(values[index.qr_token]) === token) {
      return { sheet: sheet, values: values, index: index, rowNumber: cachedRow };
    }
    cache.remove(key);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const tokens = sheet.getRange(2, index.qr_token + 1, lastRow - 1, 1).getValues();
  const entries = {};
  let rowNumber = 0;
  for (let r = 0; r < tokens.length; r++) {
    const t = cleanText_(tokens[r][0]);
    if (!t) continue;
    entries['tok:' + ssid + ':' + t] = String(r + 2);
    if (t === token) rowNumber = r + 2;
  }
  try { cache.putAll(entries, 21600); } catch (err) {}
  if (!rowNumber) return null;
  return { sheet: sheet, values: sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0], index: index, rowNumber: rowNumber };
}

function rowToRecord_(row, i) {
  return {
    registrationId: cleanText_(row[i.registration_id]), fullName: cleanText_(row[i.full_name]),
    status: cleanText_(row[i.status]), mealSelected: cleanText_(row[i.meal_selected]),
    dietaryNotes: cleanText_(row[i.dietary_notes]), mealOrdered: bool_(row[i.meal_ordered]),
    checkedInAt: row[i.checked_in_at] || '', mealRedeemedAt: row[i.meal_redeemed_at] || ''
  };
}

function checkinsSheet_(event) {
  return ensureSheet_(eventDb_(event), SHEETS.CHECKINS, CHECKIN_HEADERS);
}

function stationCheckin_(event, stationId, registrationId) {
  const cache = CacheService.getScriptCache();
  const key = 'dup:' + event.spreadsheetId + ':' + stationId + ':' + registrationId;
  const hit = cache.get(key);
  if (hit) return hit;
  const values = checkinsSheet_(event).getDataRange().getValues();
  if (values.length < 2) return '';
  const idx = headerIndex_(values[0]);
  for (let r = 1; r < values.length; r++) {
    if (cleanText_(values[r][idx.station_id]) === stationId &&
        cleanText_(values[r][idx.registration_id]) === registrationId) {
      const at = String(values[r][idx.timestamp]);
      cache.put(key, at, 21600);
      return at;
    }
  }
  return '';
}

function logScan_(event, record, mode, decision, timestamp) {
  eventDb_(event).getSheetByName(SHEETS.SCAN_LOG).appendRow([
    timestamp, event.eventId, record ? record.registrationId : '', mode,
    decision.code, Session.getActiveUser().getEmail() || 'PIN operator', decision.message
  ]);
}

// Meal scans touch two cells instead of recomputing the whole summary.
function bumpFoodSummary_(event, meal) {
  try {
    const sheet = eventDb_(event).getSheetByName(SHEETS.FOOD_SUMMARY);
    if (!sheet || sheet.getLastRow() < 2) { refreshFoodSummaryFor_(event); return; }
    const values = sheet.getDataRange().getValues();
    const idx = headerIndex_(values[0]);
    for (let r = 1; r < values.length; r++) {
      if (cleanText_(values[r][0]) === cleanText_(meal)) {
        const redeemed = Number(values[r][idx.redeemed] || 0) + 1;
        sheet.getRange(r + 1, idx.redeemed + 1).setValue(redeemed);
        if (idx.remaining !== undefined) {
          sheet.getRange(r + 1, idx.remaining + 1).setValue(Number(values[r][idx.marked_ordered] || 0) - redeemed);
        }
        return;
      }
    }
    refreshFoodSummaryFor_(event);
  } catch (err) {
    console.warn('Food summary update failed: ' + err);
  }
}

function refreshFoodSummaryFor_(event) {
  const ss = eventDb_(event);
  const reg = ss.getSheetByName(SHEETS.REGISTRATIONS).getDataRange().getValues();
  const summary = {};
  activeMealsFor_(event).forEach(function(m) { summary[m] = [m, 0, 0, 0]; });
  if (reg.length > 1) {
    const i = headerIndex_(reg[0]);
    reg.slice(1).forEach(function(r) {
      if (r[i.status] !== APP.statuses.ACTIVE) return;
      const meal = cleanText_(r[i.meal_selected]);
      if (!summary[meal]) summary[meal] = [meal, 0, 0, 0];
      summary[meal][1]++;
      if (bool_(r[i.meal_ordered])) summary[meal][2]++;
      if (r[i.meal_redeemed_at]) summary[meal][3]++;
    });
  }
  const sheet = ss.getSheetByName(SHEETS.FOOD_SUMMARY);
  sheet.clearContents();
  const rows = [['meal_name', 'requested', 'marked_ordered', 'redeemed', 'remaining']]
    .concat(Object.keys(summary).sort().map(function(k) {
      const row = summary[k];
      return row.concat([row[2] - row[3]]);
    }));
  sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#d9ead3');
}
