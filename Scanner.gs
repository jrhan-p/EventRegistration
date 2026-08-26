function processScan(scanValue, pin, mode) {
  const token = extractToken_(scanValue);
  const validMode = mode === APP.modes.CHECKIN || mode === APP.modes.MEAL;
  if (!validMode || !authorizePin_(pin, mode)) return { ok: false, code: 'UNAUTHORIZED', message: 'Operator PIN is incorrect.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findRegistration_(token);
    const record = found && rowToRecord_(found.values, found.index);
    const decision = evaluateScan_(record, mode);
    const now = new Date();
    if (decision.ok) {
      const column = mode === APP.modes.CHECKIN ? found.index.checked_in_at : found.index.meal_redeemed_at;
      found.sheet.getRange(found.rowNumber, column + 1).setValue(now);
      found.sheet.getRange(found.rowNumber, found.index.updated_at + 1).setValue(now);
      if (mode === APP.modes.MEAL) refreshFoodSummary();
    }
    logScan_(record, mode, decision, now);
    return Object.assign({}, decision, record ? {
      registrationId: record.registrationId, fullName: record.fullName,
      mealSelected: record.mealSelected, dietaryNotes: record.dietaryNotes,
      timestamp: now.toISOString()
    } : {});
  } finally {
    lock.releaseLock();
  }
}

function findRegistration_(token) {
  const sheet = db_().getSheetByName(SHEETS.REGISTRATIONS);
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

function logScan_(record, mode, decision, timestamp) {
  db_().getSheetByName(SHEETS.SCAN_LOG).appendRow([
    timestamp, setting_('EVENT_ID'), record ? record.registrationId : '', mode,
    decision.code, Session.getActiveUser().getEmail() || 'PIN operator', decision.message
  ]);
}

function refreshFoodSummary() {
  const reg = db_().getSheetByName(SHEETS.REGISTRATIONS).getDataRange().getValues();
  const summary = {};
  activeMeals_().forEach(function(m) { summary[m] = [m, 0, 0, 0]; });
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
  const sheet = db_().getSheetByName(SHEETS.FOOD_SUMMARY);
  sheet.clearContents();
  const rows = [['meal_name', 'requested', 'marked_ordered', 'redeemed']].concat(Object.keys(summary).sort().map(function(k) { return summary[k]; }));
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d9ead3');
}
