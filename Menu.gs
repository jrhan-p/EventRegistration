// Menu for whoever holds the SCRIPT PROJECT — not the spreadsheet.
//
// Worth being precise about, because the obvious assumption is wrong: this is
// a standalone project, so the master spreadsheet's sharing list and the
// project's are different ACLs. Google's installable-trigger restrictions say
// a standalone script's triggers only run for users with at least view access
// to the script file, and custom menus are officially a bound-script feature.
// So sharing the spreadsheet does NOT put this menu in anyone's Sheets UI, and
// clicking an item runs project code as the clicker, which needs project
// access too.
//
// That access cannot be granted narrowly: Script Properties are project-wide,
// so anyone who can open the project reads TWILIO_AUTH_TOKEN and GW_SA_KEY,
// the Google Wallet signing key. Add someone here only if you would also hand
// them those. For everyone else, the admin console's emailed reset code is the
// recovery path — it proves control of an organizer mailbox and grants nothing
// else.
// SpreadsheetApp.getUi() only exists while a spreadsheet is actually open, so
// every function below can ONLY be run from the RCCC Admin menu. Pressing Run
// on one in the editor throws "Cannot call SpreadsheetApp.getUi() from this
// context" — say so plainly instead, and point at the function that does work
// from the editor.
function ui_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (err) {
    console.log('This one is a menu command — open the master spreadsheet and use its RCCC Admin menu.\n' +
      'From the editor, run installAdminMenu() to put the menu there, or resetAdminPinFromEditor() to set a new admin PIN right now.');
    return null;
  }
}

function buildAdminMenu() {
  SpreadsheetApp.getUi().createMenu('RCCC Admin')
    .addItem('Set the admin PIN…', 'menuSetAdminPin')
    .addItem('Generate a random admin PIN', 'menuGenerateAdminPin')
    .addSeparator()
    .addItem('Show every station PIN', 'menuShowStationPins')
    .addItem('Open the admin console', 'menuOpenConsole')
    .addToUi();
}

// Standalone scripts get no simple onOpen, so the menu needs an installable
// trigger on the master spreadsheet. Idempotent, and re-run by setupApplication.
function installAdminMenu_(ss) {
  const target = ss || db_();
  const id = target.getId();
  const exists = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'buildAdminMenu' && t.getTriggerSourceId() === id;
  });
  if (!exists) ScriptApp.newTrigger('buildAdminMenu').forSpreadsheet(target).onOpen().create();
}

// Runnable from the script editor if the menu ever goes missing.
function installAdminMenu() {
  installAdminMenu_(null);
  return 'The RCCC Admin menu appears the next time the master spreadsheet is opened.';
}

// One validation path for every way the admin PIN can be set.
function applyNewAdminPin_(newPin) {
  const clean = cleanText_(newPin);
  // Six, not four: this PIN is checked by an endpoint anyone can reach, and
  // four digits is ten thousand guesses.
  if (!/^\d{6,10}$/.test(clean)) throw new Error('The admin PIN must be 6 to 10 digits.');
  // Deliberately NOT the strict read. Setting a PIN is the recovery path, and
  // one trashed event spreadsheet must never be able to block every route
  // back into the system; a station-PIN clash is a nuisance, being locked out
  // is not recoverable.
  if (allStationPins_()[clean]) throw new Error('That PIN already belongs to a station — pick a different one.');
  setAdminPin(clean);
  return clean;
}

function menuSetAdminPin() {
  const ui = ui_();
  if (!ui) return;
  const res = ui.prompt('Set the admin PIN', 'Type a new admin PIN (4 to 10 digits).', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  try {
    const pin = applyNewAdminPin_(res.getResponseText());
    ui.alert('Admin PIN changed',
      'The admin console now opens with ' + pin + '.\n\nAnyone still using the old PIN needs to be told the new one.',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('PIN not changed', String((err && err.message) || err), ui.ButtonSet.OK);
  }
}

function menuGenerateAdminPin() {
  const ui = ui_();
  if (!ui) return;
  const taken = allStationPins_();
  let pin;
  do { pin = String(Math.floor(Math.random() * 900000) + 100000); } while (taken[pin]);
  const res = ui.alert('New admin PIN: ' + pin,
    'Write this down now — it is not shown again and it is not stored anywhere readable.\n\nUse it as the admin PIN?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  applyNewAdminPin_(pin);
  ui.alert('Admin PIN changed', 'The admin console now opens with ' + pin + '.', ui.ButtonSet.OK);
}

function menuShowStationPins() {
  const lines = [];
  allEvents_().forEach(function(event) {
    const stations = allStationRows_(event);
    if (!stations.length) return;
    lines.push(event.name + '  [' + event.status + ']');
    stations.forEach(function(s) {
      lines.push('    ' + s.pin + '   ' + s.name +
        (s.type === 'meal' ? '  (meal pickup)' : '  (check-in)') +
        (s.active ? '' : '  — switched off'));
    });
    lines.push('');
  });
  const ui = ui_();
  if (!ui) { console.log(lines.join('\n') || 'No stations yet — create an event first.'); return; }
  ui.alert('Station PINs',
    lines.length ? lines.join('\n') + '\nOnly an ACTIVE event’s PINs open the scanner.'
      : 'No stations yet — create an event first.',
    ui.ButtonSet.OK);
}

function menuOpenConsole() {
  const ui = ui_();
  if (!ui) return;
  const url = setting_('WEB_APP_URL');
  if (!url) { ui.alert('The web app URL has not been recorded yet — open the console once from the deployment page.'); return; }
  ui.showModalDialog(
    HtmlService.createHtmlOutput('<p style="font-family:Arial,sans-serif;font-size:15px">' +
      '<a href="' + url + '?view=admin" target="_blank">Open the admin console</a></p>').setWidth(300).setHeight(80),
    'Admin console');
}

// The editor has no spreadsheet UI, so this is the one that works from the Run
// button: it invents a PIN, applies it, and prints it to the execution log.
// That log is visible to anyone with project access — the same people who
// could reset the PIN anyway.
function resetAdminPinFromEditor() {
  const taken = allStationPins_();
  let pin;
  do { pin = String(Math.floor(Math.random() * 900000) + 100000); } while (taken[pin]);
  applyNewAdminPin_(pin);
  console.log('New admin PIN: ' + pin);
  console.log('Admin console: ' + setting_('WEB_APP_URL') + '?view=admin');
  return pin;
}
