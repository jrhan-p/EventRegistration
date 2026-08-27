function processScan(scanValue, pin, mode, eventId, meetingId) {
  const token = extractToken_(scanValue);
  const validMode = mode === APP.modes.CHECKIN || mode === APP.modes.MEAL;
  if (!validMode || !authorizePin_(pin, mode)) return { ok: false, code: 'UNAUTHORIZED', message: 'Operator PIN is incorrect.' };
  const event = eventById_(eventId);
  if (!event) return { ok: false, code: 'NO_EVENT', message: 'Unknown event. Reopen the scanner from its link in the admin console.' };
  const meeting = mode === APP.modes.CHECKIN ? (cleanText_(meetingId) || 'MAIN') : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findRegistration_(token, event);
    const record = found && rowToRecord_(found.values, found.index);
    // Check-in duplicates are per meeting: swap in this meeting's prior time
    // so evaluateScan_ judges the right scope.
    if (record && mode === APP.modes.CHECKIN) record.checkedInAt = meetingCheckin_(event, meeting, record.registrationId);
    const decision = evaluateScan_(record, mode);
    const now = new Date();
    if (decision.ok) {
      if (mode === APP.modes.CHECKIN) {
        checkinsSheet_(event).appendRow([now, meeting, record.registrationId, Session.getActiveUser().getEmail() || 'PIN operator']);
        if (!found.values[found.index.checked_in_at]) found.sheet.getRange(found.rowNumber, found.index.checked_in_at + 1).setValue(now);
      } else {
        found.sheet.getRange(found.rowNumber, found.index.meal_redeemed_at + 1).setValue(now);
        refreshFoodSummaryFor_(event);
      }
      found.sheet.getRange(found.rowNumber, found.index.updated_at + 1).setValue(now);
    }
    logScan_(event, record, meeting ? mode + '@' + meeting : mode, decision, now);
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

function meetingCheckin_(event, meetingId, registrationId) {
  const values = checkinsSheet_(event).getDataRange().getValues();
  if (values.length < 2) return '';
  const idx = headerIndex_(values[0]);
  for (let r = 1; r < values.length; r++) {
    if (cleanText_(values[r][idx.meeting_id]) === meetingId &&
        cleanText_(values[r][idx.registration_id]) === registrationId) {
      return values[r][idx.timestamp];
    }
  }
  return '';
}

function findRegistration_(token, event) {
  const sheet = eventDb_(event).getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet) return null;
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const index = headerIndex_(values[0]);
  for (let r = 1; r < values.length; r++) {
    if (cleanText_(values[r][index.qr_token]) === token) return { sheet: sheet, values: values[r], index: index, rowNumber: r + 1 };
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

function authorizePin_(pin, mode) {
  const key = mode === APP.modes.CHECKIN ? 'CHECKIN_PIN_HASH' : 'MEAL_PIN_HASH';
  const expected = PropertiesService.getScriptProperties().getProperty(key);
  return Boolean(expected) && hash_(String(pin)) === expected;
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
