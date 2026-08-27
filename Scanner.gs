// The PIN alone routes the scan: it identifies the event, the station, and
// whether this is a check-in desk or the meal window.
function processScan(scanValue, pin) {
  const token = extractToken_(scanValue);
  const station = resolvePin_(pin);
  if (!station) return { ok: false, code: 'UNAUTHORIZED', message: 'Operator PIN is incorrect.' };
  const event = station.event;
  const mode = station.type === 'meal' ? APP.modes.MEAL : APP.modes.CHECKIN;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findRegistration_(token, event);
    const record = found && rowToRecord_(found.values, found.index);
    // Check-in duplicates are per station: swap in this station's prior time
    // so evaluateScan_ judges the right scope.
    if (record && mode === APP.modes.CHECKIN) record.checkedInAt = stationCheckin_(event, station.stationId, record.registrationId);
    const decision = evaluateScan_(record, mode);
    const now = new Date();
    if (decision.ok) {
      if (mode === APP.modes.CHECKIN) {
        checkinsSheet_(event).appendRow([now, station.stationId, record.registrationId, Session.getActiveUser().getEmail() || 'PIN operator']);
        if (!found.values[found.index.checked_in_at]) found.sheet.getRange(found.rowNumber, found.index.checked_in_at + 1).setValue(now);
      } else {
        found.sheet.getRange(found.rowNumber, found.index.meal_redeemed_at + 1).setValue(now);
        refreshFoodSummaryFor_(event);
      }
      found.sheet.getRange(found.rowNumber, found.index.updated_at + 1).setValue(now);
    }
    logScan_(event, record, mode + '@' + station.stationId, decision, now);
    // The check-in desk must not see meal choices; only the kitchen scanner may.
    const extra = record ? {
      registrationId: record.registrationId, fullName: record.fullName, timestamp: now.toISOString()
    } : {};
    if (record && mode === APP.modes.MEAL) {
      extra.mealSelected = record.mealSelected;
      extra.dietaryNotes = record.dietaryNotes;
    }
    return Object.assign({}, decision, extra);
  } finally {
    lock.releaseLock();
  }
}

function checkinsSheet_(event) {
  return ensureSheet_(eventDb_(event), SHEETS.CHECKINS, CHECKIN_HEADERS);
}

function stationCheckin_(event, stationId, registrationId) {
  const values = checkinsSheet_(event).getDataRange().getValues();
  if (values.length < 2) return '';
  const idx = headerIndex_(values[0]);
  for (let r = 1; r < values.length; r++) {
    if (cleanText_(values[r][idx.station_id]) === stationId &&
        cleanText_(values[r][idx.registration_id]) === registrationId) {
      return values[r][idx.timestamp];
    }
  }
  return '';
}

// Reads only the token column, then the single matching row — the whole-sheet
// read was the scan path's biggest payload.
function findRegistration_(token, event) {
  const sheet = eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;
  const index = headerIndex_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  const tokens = sheet.getRange(2, index.qr_token + 1, lastRow - 1, 1).getValues();
  for (let r = 0; r < tokens.length; r++) {
    if (cleanText_(tokens[r][0]) === token) {
      const rowNumber = r + 2;
      return { sheet: sheet, values: sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0], index: index, rowNumber: rowNumber };
    }
  }
  return null;
}

function rowToRecord_(row, i) {
  return {
    registrationId: cleanText_(row[i.registration_id]), fullName: cleanText_(row[i.full_name]),
    status: cleanText_(row[i.status]), mealSelected: cleanText_(row[i.meal_selected]),
    dietaryNotes: cleanText_(row[i.dietary_notes]), mealOrdered: bool_(row[i.meal_ordered]),
    checkedInAt: row[i.checked_in_at] || '', mealRedeemedAt: row[i.meal_redeemed_at] || ''
  };
}

function logScan_(event, record, mode, decision, timestamp) {
  eventDb_(event).getSheetByName(SHEETS.SCAN_LOG).appendRow([
    timestamp, event.eventId, record ? record.registrationId : '', mode,
    decision.code, Session.getActiveUser().getEmail() || 'PIN operator', decision.message
  ]);
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
  const rows = [['meal_name', 'requested', 'marked_ordered', 'redeemed']].concat(Object.keys(summary).sort().map(function(k) { return summary[k]; }));
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d9ead3');
}
