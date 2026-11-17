/**
 * Friends Circle Bachat Gat - savings group app.
 *
 * A Google Apps Script web app that lives inside the "Savings Group Ledger"
 * Google Sheet. The sheet stays the only database:
 *   Members           A:H  (Member ID, Name, Phone, Joined, Status, Opening savings, Opening loan, Notes)
 *   Monthly entries   A:K  (Month, Member ID, [name formula], Deposit, Loan taken, Loan repaid,
 *                           Interest paid, Penalty, [total formula], Notes, Date paid)
 *                           Column K "Date paid" is added by the app the first time it saves.
 *   Cash handed over  A:D  (Month, Handed to, Amount, Notes)
 *   Summary           B5:B7 group rules (deposit, interest rate, rounding)
 *
 * The app never writes to columns C and I of "Monthly entries" because those
 * hold ARRAYFORMULAs.
 *
 * Protection:
 *   - Members log in with their member number and a 4 digit PIN.
 *   - The treasurer logs in as "ADMIN" with a PIN of at least 6 digits.
 *   - PINs are stored only as salted SHA-256 hashes in Script Properties,
 *     never in the sheet.
 *   - 5 wrong PINs lock that login for 15 minutes.
 *   - A member sees their own monthly entries, and every member's savings and
 *     loan balance (the group wants full transparency on loans).
 *   - Only the ADMIN login can write to the sheet.
 */

var SHEET_MEMBERS = 'Members';
var SHEET_ENTRIES = 'Monthly entries';
var SHEET_CASH = 'Cash handed over';
var SHEET_SUMMARY = 'Summary';

var ADMIN_ID = 'ADMIN';
var SESSION_SECONDS = 6 * 60 * 60;   // CacheService maximum
var MAX_FAILS = 5;
var LOCK_SECONDS = 15 * 60;
var GROUP_NAME = 'Friends Circle Bachat Gat';

/* ------------------------------------------------------------------ */
/* Web app entry point and sheet menu                                  */
/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(GROUP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    // Lets the app open inside the group's own app page (GitHub Pages) with its icon and full screen.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Savings App')
    .addItem('1. Set treasurer (ADMIN) PIN', 'menuSetAdminPin')
    .addItem('2. Make PINs for all members', 'menuGenerateAllPins')
    .addItem('Reset one member\'s PIN', 'menuResetOnePin')
    .addItem('Unlock a locked login', 'menuUnlock')
    .addSeparator()
    .addItem('3. Turn on 9th-of-month reminder', 'menuTurnOnReminder')
    .addItem('Turn off 9th-of-month reminder', 'menuTurnOffReminder')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* Public functions called from the app (google.script.run)            */
/* ------------------------------------------------------------------ */

function apiLogin(loginId, pin) {
  var id = normaliseLoginId_(loginId);
  if (!id) return fail_('bad_login');
  pin = String(pin || '').trim();

  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + id;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) return fail_('locked');

  var stored = PropertiesService.getScriptProperties().getProperty(pinKey_(id));
  if (!stored || !checkPin_(pin, stored)) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    return fail_(fails + 1 >= MAX_FAILS ? 'locked' : 'bad_login');
  }
  cache.remove(failKey);

  var role = id === ADMIN_ID ? 'admin' : 'member';
  if (role === 'member') {
    var member = getMembers_().filter(function (m) { return m.id === id; })[0];
    if (!member) return fail_('bad_login');
  }
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('sess_' + token, JSON.stringify({ id: id, role: role }), SESSION_SECONDS);
  return { ok: true, token: token, role: role, id: id };
}

function apiLogout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return { ok: true };
}

function apiChangePin(token, oldPin, newPin) {
  var s = session_(token);
  if (!s) return fail_('session');
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty(pinKey_(s.id));
  if (!stored || !checkPin_(String(oldPin || ''), stored)) return fail_('wrong_old_pin');
  var err = validatePin_(s.id, String(newPin || ''));
  if (err) return fail_(err);
  props.setProperty(pinKey_(s.id), hashPin_(String(newPin)));
  return { ok: true };
}

/** Everything a member is allowed to see: their own numbers and group totals. */
function apiMyData(token) {
  var s = session_(token);
  if (!s) return fail_('session');
  var book = loadBook_();
  var result = { ok: true, group: groupTotals_(book), rules: book.rules, groupName: GROUP_NAME };
  if (s.role === 'member') {
    var me = memberBalance_(book, s.id);
    result.me = me;
    result.history = book.entries
      .filter(function (e) { return e.id === s.id; })
      .sort(function (a, b) { return a.month < b.month ? 1 : -1; })
      .map(function (e) {
        return { month: e.month, deposit: e.deposit, loanTaken: e.loanTaken, repaid: e.repaid,
                 interest: e.interest, penalty: e.penalty, note: e.note, paidDate: e.paidDate };
      });
  }
  // Transparency: every member can see everyone's savings and loans.
  result.members = book.members.map(function (m) {
    var b = memberBalance_(book, m.id);
    return { id: b.id, name: b.name, status: b.status, savings: b.savings, loan: b.loan,
             interestDue: b.interestDue, net: b.net };
  });
  result.role = s.role;
  return result;
}

/** Any logged-in member: who paid what, and on which date, in one month. */
function apiMonthView(token, month) {
  var s = session_(token);
  if (!s) return fail_('session');
  month = normaliseMonth_(month);
  if (!month) return fail_('bad_month');
  var book = loadBook_();
  var rows = book.members.map(function (m) {
    var ex = book.entries.filter(function (e) { return e.id === m.id && e.month === month; });
    return {
      id: m.id, name: m.name, status: m.status, hasRow: ex.length > 0,
      deposit: sum_(ex, 'deposit'), interest: sum_(ex, 'interest'), repaid: sum_(ex, 'repaid'),
      penalty: sum_(ex, 'penalty'), loanTaken: sum_(ex, 'loanTaken'),
      paidDate: ex.map(function (e) { return e.paidDate; }).filter(String)[0] || ''
    };
  }).filter(function (r) { return r.status === 'active' || r.hasRow; });
  var months = {};
  book.entries.forEach(function (e) { if (e.month) months[e.month] = true; });
  return { ok: true, month: month, rows: rows, months: Object.keys(months).sort().reverse(),
           cash: sum_(book.cash.filter(function (c) { return c.month === month; }), 'amount') };
}

/** Treasurer: what each member owes this month, for the WhatsApp reminders. */
function apiReminders(token, month) {
  var s = session_(token);
  if (!s || s.role !== 'admin') return fail_('session');
  month = normaliseMonth_(month);
  if (!month) return fail_('bad_month');
  var book = loadBook_();
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var rows = book.members.filter(function (m) { return m.status === 'active'; }).map(function (m) {
    var b = memberBalance_(book, m.id);
    var loanStart = loanOutstanding_(book, m.id, month);
    var ex = book.entries.filter(function (e) { return e.id === m.id && e.month === month; });
    return {
      id: m.id, name: m.name, phone: m.phone,
      depositDue: book.rules.deposit, interestDue: interestFor_(loanStart, book.rules),
      savings: b.savings, loan: b.loan,
      paid: ex.some(function (e) { return e.deposit > 0; })
    };
  });
  return { ok: true, month: month, rows: rows, url: url, groupName: GROUP_NAME };
}

/** Treasurer: every member's balance (for checking). */
function apiAllBalances(token) {
  var s = session_(token);
  if (!s || s.role !== 'admin') return fail_('session');
  var book = loadBook_();
  return {
    ok: true,
    group: groupTotals_(book),
    members: book.members.map(function (m) { return memberBalance_(book, m.id); })
  };
}

/** Treasurer: one month's entries, with suggested deposit and interest. */
function apiGetMonth(token, month) {
  var s = session_(token);
  if (!s || s.role !== 'admin') return fail_('session');
  month = normaliseMonth_(month);
  if (!month) return fail_('bad_month');
  var book = loadBook_();
  var rows = book.members
    .filter(function (m) { return m.status === 'active' || book.entries.some(function (e) { return e.id === m.id && e.month === month; }); })
    .map(function (m) {
      var existing = book.entries.filter(function (e) { return e.id === m.id && e.month === month; });
      var loanBefore = loanOutstanding_(book, m.id, month);
      var row = {
        id: m.id, name: m.name, loanBefore: loanBefore,
        suggestedInterest: interestFor_(loanBefore, book.rules),
        suggestedDeposit: book.rules.deposit,
        saved: existing.length > 0,
        deposit: sum_(existing, 'deposit'), loanTaken: sum_(existing, 'loanTaken'),
        repaid: sum_(existing, 'repaid'), interest: sum_(existing, 'interest'),
        penalty: sum_(existing, 'penalty'),
        note: existing.map(function (e) { return e.note; }).filter(String).join('; '),
        paidDate: existing.map(function (e) { return e.paidDate; }).filter(String)[0] || '',
        duplicateRows: existing.length > 1
      };
      return row;
    });
  var cash = book.cash.filter(function (c) { return c.month === month; });
  return { ok: true, month: month, rows: rows, cash: cash, rules: book.rules };
}

/**
 * Treasurer: save one month. `rows` is a list of
 * {id, deposit, loanTaken, repaid, interest, penalty, note}.
 * Updates the member's existing row for that month, or adds a new row.
 */
function apiSaveMonth(token, month, rows) {
  var s = session_(token);
  if (!s || s.role !== 'admin') return fail_('session');
  month = normaliseMonth_(month);
  if (!month) return fail_('bad_month');
  if (!rows || !rows.length) return fail_('nothing');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return fail_('busy');
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_ENTRIES);
    var tz = ss.getSpreadsheetTimeZone();
    var memberIds = {};
    getMembers_().forEach(function (m) { memberIds[m.id] = true; });

    var last = lastDataRow_(sh);
    var data = last >= 2 ? sh.getRange(2, 1, last - 1, 11).getValues() : [];
    var index = {};
    data.forEach(function (r, i) {
      var key = monthKey_(r[0], tz) + '|' + String(r[1]).trim().toUpperCase();
      if (index[key] === undefined) index[key] = i + 2;
    });

    var updated = 0, added = 0, newRows = [];
    rows.forEach(function (r) {
      var id = String(r.id || '').trim().toUpperCase();
      if (!memberIds[id]) throw new Error('Unknown member ' + id);
      var nums = ['deposit', 'loanTaken', 'repaid', 'interest', 'penalty'].map(function (k) {
        var v = Number(r[k] || 0);
        if (!isFinite(v) || v < 0 || v > 10000000) throw new Error('Bad amount for ' + id);
        return v === 0 ? '' : Math.round(v * 100) / 100;
      });
      var note = String(r.note || '').slice(0, 300);
      var paid = parseDay_(r.paidDate);
      var rowNo = index[month + '|' + id];
      if (rowNo) {
        sh.getRange(rowNo, 4, 1, 5).setValues([nums]);
        if (note) sh.getRange(rowNo, 10).setValue(note);
        sh.getRange(rowNo, 11).setValue(paid || '');
        updated++;
      } else if (nums.some(function (v) { return v !== ''; })) {
        newRows.push({ id: id, nums: nums, note: note, paid: paid || '' });
      }
    });

    if (newRows.length) {
      var start = last + 1;
      var parts = month.split('-');
      var monthDate = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
      sh.getRange(start, 1, newRows.length, 2)
        .setValues(newRows.map(function (n) { return [monthDate, n.id]; }));
      sh.getRange(start, 1, newRows.length, 1).setNumberFormat('mmm yyyy');
      sh.getRange(start, 4, newRows.length, 5)
        .setValues(newRows.map(function (n) { return n.nums; }));
      sh.getRange(start, 10, newRows.length, 1)
        .setValues(newRows.map(function (n) { return [n.note]; }));
      sh.getRange(start, 11, newRows.length, 1)
        .setValues(newRows.map(function (n) { return [n.paid]; }));
      added = newRows.length;
    }
    if (updated || added) {
      var head = sh.getRange(1, 11);
      if (String(head.getValue()).trim() === '') head.setValue('Date paid');
      sh.getRange(2, 11, Math.max(lastDataRow_(sh) - 1, 1), 1).setNumberFormat('d mmm yyyy');
    }
    SpreadsheetApp.flush();
    return { ok: true, updated: updated, added: added };
  } catch (e) {
    return { ok: false, error: 'save_failed', detail: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

/** Treasurer: record cash handed over to someone for safe keeping. */
function apiAddCash(token, month, handedTo, amount, note) {
  var s = session_(token);
  if (!s || s.role !== 'admin') return fail_('session');
  month = normaliseMonth_(month);
  amount = Number(amount);
  if (!month || !handedTo || !(amount > 0)) return fail_('bad_input');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return fail_('busy');
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CASH);
    var row = lastDataRow_(sh, 1) + 1;
    var parts = month.split('-');
    sh.getRange(row, 1, 1, 4).setValues([[new Date(Number(parts[0]), Number(parts[1]) - 1, 1),
      String(handedTo).slice(0, 100), amount, String(note || '').slice(0, 300)]]);
    sh.getRange(row, 1).setNumberFormat('mmm yyyy');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Sheet menu actions (only people who can edit the sheet see these)   */
/* ------------------------------------------------------------------ */

function menuSetAdminPin() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Treasurer PIN', 'Type a new PIN for the ADMIN login (at least 6 digits):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pin = r.getResponseText().trim();
  var err = validatePin_(ADMIN_ID, pin);
  if (err) { ui.alert('PIN not saved: it must be at least 6 digits.'); return; }
  PropertiesService.getScriptProperties().setProperty(pinKey_(ADMIN_ID), hashPin_(pin));
  ui.alert('Treasurer PIN saved. Log in to the app with ID "ADMIN" and this PIN.');
}

function menuGenerateAllPins() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var members = getMembers_();
  var missing = members.filter(function (m) { return !props.getProperty(pinKey_(m.id)); });
  var target = missing;
  if (!missing.length) {
    var a = ui.alert('Every member already has a PIN. Make NEW PINs for everyone? Old PINs will stop working.', ui.ButtonSet.YES_NO);
    if (a !== ui.Button.YES) return;
    target = members;
  }
  var lines = [];
  target.forEach(function (m) {
    var pin = randomPin_();
    props.setProperty(pinKey_(m.id), hashPin_(pin));
    lines.push('<tr><td>' + esc_(m.id.replace(/^M0?/, '')) + '</td><td>' + esc_(m.name) +
               '</td><td style="font-size:18px"><b>' + pin + '</b></td></tr>');
  });
  showPinSlip_(lines, 'New PINs (shown only once)');
}

function menuResetOnePin() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Reset PIN', 'Member number or ID (for example 5 or M05):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var id = normaliseLoginId_(r.getResponseText());
  var m = getMembers_().filter(function (x) { return x.id === id; })[0];
  if (!m) { ui.alert('No member with that number.'); return; }
  var pin = randomPin_();
  PropertiesService.getScriptProperties().setProperty(pinKey_(m.id), hashPin_(pin));
  CacheService.getScriptCache().remove('fail_' + m.id);
  showPinSlip_(['<tr><td>' + esc_(m.id.replace(/^M0?/, '')) + '</td><td>' + esc_(m.name) +
                '</td><td style="font-size:18px"><b>' + pin + '</b></td></tr>'], 'New PIN for ' + m.name);
}

function menuUnlock() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Unlock', 'Member number, member ID, or ADMIN:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var id = normaliseLoginId_(r.getResponseText());
  if (!id) { ui.alert('Not a valid login.'); return; }
  CacheService.getScriptCache().remove('fail_' + id);
  ui.alert(id + ' is unlocked.');
}

function menuTurnOnReminder() {
  removeReminderTriggers_();
  ScriptApp.newTrigger('sendMonthlyReminder').timeBased().onMonthDay(9).atHour(8).create();
  SpreadsheetApp.getUi().alert('Done. On the 9th of every month, around 8-9 AM, you will get an email with the list of members and a link to send each one their WhatsApp message.');
}

function menuTurnOffReminder() {
  var n = removeReminderTriggers_();
  SpreadsheetApp.getUi().alert(n ? 'The 9th-of-month reminder is turned off.' : 'The reminder was not on.');
}

function removeReminderTriggers_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'sendMonthlyReminder') { ScriptApp.deleteTrigger(tr); n++; }
  });
  return n;
}

/** Runs by itself on the 9th: emails the treasurer (the sheet owner) the month's dues. */
function sendMonthlyReminder() {
  var ss = SpreadsheetApp.getActive();
  var month = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM');
  var book = loadBook_();
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var lines = book.members.filter(function (m) { return m.status === 'active'; }).map(function (m) {
    var due = book.rules.deposit + interestFor_(loanOutstanding_(book, m.id, month), book.rules);
    var paid = book.entries.some(function (e) { return e.id === m.id && e.month === month && e.deposit > 0; });
    return (paid ? '[paid] ' : '[due]  ') + m.id + ' ' + m.name + ': Rs ' + due;
  });
  var to = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(to, GROUP_NAME + ': send this month\'s reminders',
    'Today is the 9th. Open the app, log in as ADMIN and go to the Reminders page to send each member their WhatsApp message:\n' +
    url + '\n\nThis month (' + month + '):\n' + lines.join('\n'));
}

function showPinSlip_(rows, title) {
  var html = '<div style="font-family:sans-serif">' +
    '<p>Write these down or print this box now. They are not stored anywhere readable and will not be shown again.</p>' +
    '<table border="1" cellpadding="6" style="border-collapse:collapse">' +
    '<tr><th>Member no.</th><th>Name</th><th>PIN</th></tr>' + rows.join('') + '</table>' +
    '<p><button onclick="window.print()">Print</button></p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(520), title);
}

/* ------------------------------------------------------------------ */
/* Ledger maths (mirrors the formulas on the Summary tab)              */
/* ------------------------------------------------------------------ */

function loadBook_() {
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone();
  var members = getMembers_();

  var es = ss.getSheetByName(SHEET_ENTRIES);
  var eLast = lastDataRow_(es);
  var entries = eLast >= 2 ? es.getRange(2, 1, eLast - 1, 11).getValues() : [];
  entries = entries
    .filter(function (r) { return String(r[1]).trim(); })
    .map(function (r) {
      return {
        month: monthKey_(r[0], tz), id: String(r[1]).trim().toUpperCase(),
        deposit: num_(r[3]), loanTaken: num_(r[4]), repaid: num_(r[5]),
        interest: num_(r[6]), penalty: num_(r[7]), note: String(r[9] || ''),
        paidDate: dayKey_(r[10], tz)
      };
    });

  var cs = ss.getSheetByName(SHEET_CASH);
  var cLast = lastDataRow_(cs, 1);
  var cash = cLast >= 2 ? cs.getRange(2, 1, cLast - 1, 4).getValues() : [];
  cash = cash.filter(function (r) { return r[0] !== '' || r[2] !== ''; })
    .map(function (r) { return { month: monthKey_(r[0], tz), to: String(r[1]), amount: num_(r[2]), note: String(r[3] || '') }; });

  return { members: members, entries: entries, cash: cash, rules: getRules_() };
}

function getMembers_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MEMBERS);
  var last = lastDataRow_(sh, 1);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      return {
        id: String(r[0]).trim().toUpperCase(), name: String(r[1]).trim(), phone: String(r[2] || '').trim(),
        status: String(r[4]).trim().toLowerCase(),
        openingSavings: num_(r[5]), openingLoan: num_(r[6])
      };
    });
}

function getRules_() {
  var vals = SpreadsheetApp.getActive().getSheetByName(SHEET_SUMMARY).getRange('A1:B10').getValues();
  var rules = { deposit: 1000, rate: 0.02, roundTo: 5 };
  vals.forEach(function (r) {
    var label = String(r[0]).toLowerCase();
    if (label.indexOf('monthly deposit per member') === 0) rules.deposit = num_(r[1]) || rules.deposit;
    else if (label.indexOf('interest per month') === 0) rules.rate = num_(r[1]);
    else if (label.indexOf('round interest up') === 0) rules.roundTo = num_(r[1]);
  });
  if (rules.rate > 1) rules.rate = rules.rate / 100; // typed as 2 instead of 2%
  return rules;
}

function memberBalance_(book, id) {
  var m = book.members.filter(function (x) { return x.id === id; })[0];
  if (!m) return null;
  var mine = book.entries.filter(function (e) { return e.id === id; });
  var deposited = m.openingSavings + sum_(mine, 'deposit');
  var g = groupTotals_(book);
  var share = (m.status === 'active' && g.activeMembers > 0) ? g.interestEarned / g.activeMembers : 0;
  var loan = m.openingLoan + sum_(mine, 'loanTaken') - sum_(mine, 'repaid');
  var savings = deposited + share;
  return {
    id: m.id, name: m.name, status: m.status,
    deposited: deposited, interestShare: round2_(share), savings: round2_(savings),
    loan: loan, interestDue: interestFor_(loan, book.rules),
    interestPaid: sum_(mine, 'interest'), penaltyPaid: sum_(mine, 'penalty'),
    net: round2_(savings - loan)
  };
}

function groupTotals_(book) {
  var e = book.entries;
  var deposits = book.members.reduce(function (s, m) { return s + m.openingSavings; }, 0) + sum_(e, 'deposit');
  var earned = sum_(e, 'interest') + sum_(e, 'penalty');
  var loans = book.members.reduce(function (s, m) { return s + m.openingLoan; }, 0) + sum_(e, 'loanTaken') - sum_(e, 'repaid');
  var paidIn = e.reduce(function (s, x) { return s + x.deposit + x.repaid + x.interest + x.penalty; }, 0);
  var cash = sum_(book.cash, 'amount');
  var months = {};
  e.forEach(function (x) { if (x.month) months[x.month] = true; });
  return {
    totalDeposits: deposits, interestEarned: earned, totalSavings: deposits + earned,
    loansOutstanding: loans, cashNotHandedOver: paidIn - sum_(e, 'loanTaken') - cash,
    activeMembers: book.members.filter(function (m) { return m.status === 'active'; }).length,
    lastMonth: Object.keys(months).sort().pop() || ''
  };
}

/** Loan balance at the START of `month` (only earlier months count). */
function loanOutstanding_(book, id, month) {
  var m = book.members.filter(function (x) { return x.id === id; })[0];
  var bal = m ? m.openingLoan : 0;
  book.entries.forEach(function (e) {
    if (e.id === id && e.month && e.month < month) bal += e.loanTaken - e.repaid;
  });
  return bal;
}

function interestFor_(loan, rules) {
  if (!(loan > 0)) return 0;
  var raw = loan * rules.rate;
  if (rules.roundTo > 0) return Math.ceil(Math.round(raw * 100) / 100 / rules.roundTo) * rules.roundTo;
  return round2_(raw);
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function session_(token) {
  if (!token || typeof token !== 'string' || token.length > 100) return null;
  var raw = CacheService.getScriptCache().get('sess_' + token);
  return raw ? JSON.parse(raw) : null;
}

function normaliseLoginId_(v) {
  var s = String(v || '').trim().toUpperCase();
  if (s === ADMIN_ID) return ADMIN_ID;
  var m = s.match(/^M?(\d{1,3})$/);
  if (!m) return null;
  var n = Number(m[1]);
  if (n < 1) return null;
  return 'M' + (n < 10 ? '0' + n : String(n));
}

function normaliseMonth_(v) {
  var s = String(v || '').trim();
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(s) ? s : null;
}

/** Turns a cell value (date or text like "Aug 2026") into "2026-08". */
function monthKey_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM');
  var s = String(v || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
  var names = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  var m = s.toLowerCase().match(/^([a-z]{3})[a-z]*\.?\s+(\d{4})$/);
  if (m && names.indexOf(m[1]) >= 0) {
    var mm = names.indexOf(m[1]) + 1;
    return m[2] + '-' + (mm < 10 ? '0' + mm : mm);
  }
  return s;
}

/** "2026-11-05" from a date cell or text. */
function dayKey_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

/** "2026-11-05" from the app -> Date, or null. */
function parseDay_(v) {
  var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d) ? null : d;
}

/** Last row that has something in column `col` (default B, the member ID). */
function lastDataRow_(sh, col) {
  col = col || 2;
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var vals = sh.getRange(1, col, last, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

function num_(v) {
  if (typeof v === 'number') return v;
  var n = Number(String(v || '').replace(/[₹,\s]/g, ''));
  return isFinite(n) ? n : 0;
}
function sum_(list, key) { return list.reduce(function (s, x) { return s + (Number(x[key]) || 0); }, 0); }
function round2_(n) { return Math.round(n * 100) / 100; }
function fail_(code) { return { ok: false, error: code }; }
function pinKey_(id) { return 'PIN_' + id; }

function validatePin_(id, pin) {
  var min = id === ADMIN_ID ? 6 : 4;
  if (!/^\d+$/.test(pin) || pin.length < min || pin.length > 12) return 'bad_new_pin';
  if (/^(\d)\1+$/.test(pin) || '0123456789012'.indexOf(pin) >= 0) return 'weak_pin';
  return null;
}

function randomPin_() {
  var pin;
  do {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    var n = ((bytes[0] & 0xff) << 8 | (bytes[1] & 0xff)) % 10000;
    pin = ('000' + n).slice(-4);
  } while (validatePin_('M01', pin));
  return pin;
}

function hashPin_(pin) {
  var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  return salt + ':' + digest_(salt + ':' + pin);
}

function checkPin_(pin, stored) {
  var parts = String(stored).split(':');
  if (parts.length !== 2) return false;
  var a = digest_(parts[0] + ':' + pin), b = parts[1];
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function digest_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
