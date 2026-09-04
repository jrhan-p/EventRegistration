const APP = Object.freeze({
  version: '0.2.0',
  qrPrefix: 'CER1:',
  noMeal: 'No meal',
  statuses: Object.freeze({ ACTIVE: 'ACTIVE', CANCELED: 'CANCELED' }),
  modes: Object.freeze({ CHECKIN: 'checkin', MEAL: 'meal' })
});

function cleanText_(value) {
  return String(value == null ? '' : value).trim();
}

function normalizePhone_(value) {
  const raw = cleanText_(value);
  const digits = raw.replace(/\D/g, '');
  if (digits.length === 10) return '+1' + digits;
  if (digits.length >= 11 && digits.length <= 15) return '+' + digits;
  return '';
}

function bool_(value) {
  if (value === true) return true;
  return /^(true|yes|y|1)$/i.test(cleanText_(value));
}

function extractToken_(scanValue) {
  const value = cleanText_(scanValue);
  if (value.indexOf(APP.qrPrefix) === 0) return value.slice(APP.qrPrefix.length);
  const match = value.match(/[?&]token=([^&#]+)/i);
  return match ? decodeURIComponent(match[1]) : value;
}

function evaluateScan_(record, mode) {
  if (!record) return { ok: false, code: 'NOT_FOUND', message: 'QR code is not recognized.' };
  if (record.status !== APP.statuses.ACTIVE) {
    return { ok: false, code: 'INACTIVE', message: 'This registration is not active.' };
  }
  if (mode === APP.modes.CHECKIN) {
    if (record.checkedInAt) {
      return { ok: false, code: 'ALREADY_CHECKED_IN', message: 'Already checked in.', priorAt: record.checkedInAt };
    }
    return { ok: true, code: 'CHECK_IN', message: 'Check-in approved.' };
  }
  if (mode === APP.modes.MEAL) {
    if (!record.mealSelected || record.mealSelected === APP.noMeal) {
      return { ok: false, code: 'NO_MEAL', message: 'No meal was selected for this attendee.' };
    }
    if (!record.mealOrdered) {
      return { ok: false, code: 'NOT_ORDERED', message: 'This meal was not marked as ordered.' };
    }
    if (record.mealRedeemedAt) {
      return { ok: false, code: 'ALREADY_REDEEMED', message: 'Meal was already redeemed.', priorAt: record.mealRedeemedAt };
    }
    return { ok: true, code: 'REDEEM_MEAL', message: 'Meal pickup approved.' };
  }
  return { ok: false, code: 'BAD_MODE', message: 'Scanner mode is invalid.' };
}
