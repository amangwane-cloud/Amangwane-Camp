/**
 * Amangwane booking backend
 * ---------------------------------------------------------------------------
 * Receives booking requests from booking.html and writes them into THIS
 * workbook (the "Amangwane Bookings" sheet).
 *
 * Install:
 *   1. Open the workbook > Extensions > Apps Script.
 *   2. Add a new file next to google-sheets-availability-setup.gs and paste
 *      this in. (Keep both files in the same project.)
 *   3. Set SHARED_SECRET below to a long random string.
 *   4. Deploy > New deployment > type "Web app":
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Copy the /exec URL it gives you.
 *   5. Put the SAME secret and that URL into booking.html (CONFIG block).
 *
 * Re-deploy (Deploy > Manage deployments > edit > new version) every time you
 * change this file, or the live URL keeps running the old code.
 * ---------------------------------------------------------------------------
 */

const TZ = 'Africa/Johannesburg';
const SHARED_SECRET = 'REPLACE_WITH_SHARED_SECRET'; // must match booking.html CONFIG.secret
const DEPOSIT_RATE = 0.5;
const DEPOSIT_WINDOW_HOURS = 48;

const TAB_REGISTER = 'Booking Register';
const TAB_ROOMS = 'Room Setup';
const TAB_BOARD = 'Availability Board';

const BOARD_HEADER_ROWS = 1; // row 1 is "Date | Day | Holiday... | Tent 1 ...  Tent 10"
const BOARD_DATE_COL = 1;    // column A
const BOARD_FIRST_UNIT_COL = 4; // column D = "Tent 1"

const OK_BOARD_STATES = ['Available']; // a night is bookable only if the cell reads exactly this

/** GET: health check + availability feed for the booking form. */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.secret !== SHARED_SECRET) return json({ ok: false, error: 'Unauthorised' });
    if (p.action === 'availability') {
      return json({ ok: true, blocked: blockedDates(p.start, p.end) });
    }
    return json({ ok: true, status: 'Amangwane booking backend live' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** POST: create a booking (reserve nights on the board + append a register row). */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // serialise submissions so two guests can't grab the same tent
  } catch (err) {
    return json({ ok: false, error: 'Server busy, please try again.' });
  }
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.secret !== SHARED_SECRET) return json({ ok: false, error: 'Unauthorised' });
    if (body.company) return json({ ok: false, error: 'Rejected' }); // honeypot field

    const req = validateRequest(body);
    if (req.error) return json({ ok: false, error: req.error });

    const ss = SpreadsheetApp.getActive();
    const board = ss.getSheetByName(TAB_BOARD);
    const unitCol = boardUnitColumn(board, req.unit);
    if (unitCol < 0) return json({ ok: false, error: 'Unknown unit: ' + req.unit });

    const rowMap = boardRowMap(board);
    const nights = dateKeys(req.checkin, req.nights);
    const rows = [];
    for (const key of nights) {
      if (!rowMap[key]) return json({ ok: false, error: 'Those dates are outside the availability calendar.' });
      rows.push(rowMap[key]);
    }
    for (const row of rows) {
      const state = String(board.getRange(row, unitCol).getValue()).trim();
      if (OK_BOARD_STATES.indexOf(state) < 0) {
        return json({ ok: false, error: req.unit + ' is no longer available for those dates.' });
      }
    }

    // reserve every night
    for (const row of rows) board.getRange(row, unitCol).setValue('Deposit pending');

    // append the register row (column order must match the Booking Register header)
    const now = new Date();
    const depositDue = new Date(now.getTime() + DEPOSIT_WINDOW_HOURS * 3600 * 1000);
    const deposit = Math.round(req.total * DEPOSIT_RATE);
    const balance = req.total - deposit;
    const bookingId = nextBookingId(ss);

    ss.getSheetByName(TAB_REGISTER).appendRow([
      bookingId, 'deposit-pending', req.firstName, req.lastName, req.email, req.phone,
      req.unit, req.unitType, req.guests, req.checkinStr, req.checkoutStr, req.nights,
      req.rate, req.total, deposit, 0,
      Utilities.formatDate(depositDue, TZ, 'yyyy-MM-dd'), balance,
      'Website', Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"), req.notes
    ]);

    return json({ ok: true, ref: bookingId, nights: req.nights, total: req.total, deposit: deposit });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------- helpers -------------------------------- */

function validateRequest(b) {
  const firstName = String(b.firstName || '').trim();
  const lastName = String(b.lastName || '').trim();
  const email = String(b.email || '').trim();
  const phone = String(b.phone || '').trim();
  const unit = String(b.unit || '').trim();
  const guests = parseInt(b.guests, 10) || 0;
  const notes = String(b.notes || '').trim().slice(0, 500);
  const checkinStr = String(b.checkin || '').slice(0, 10);
  const checkoutStr = String(b.checkout || '').slice(0, 10);

  if (!firstName || !lastName) return { error: 'First and last name are required.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { error: 'A valid email address is required.' };
  if (!phone) return { error: 'A phone number is required.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(checkinStr) || !/^\d{4}-\d{2}-\d{2}$/.test(checkoutStr)) {
    return { error: 'Valid check-in and check-out dates are required.' };
  }

  const checkin = parseISO(checkinStr);
  const checkout = parseISO(checkoutStr);
  const today = parseISO(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
  if (!(checkout > checkin)) return { error: 'Check-out must be after check-in.' };
  if (checkin < today) return { error: 'Check-in cannot be in the past.' };
  const nights = Math.round((checkout - checkin) / 86400000);

  const room = lookupRoom(unit);
  if (!room) return { error: 'Unknown unit: ' + unit };
  if (String(room.active).toLowerCase() !== 'yes') return { error: unit + ' is not open for booking.' };
  if (guests < 1) return { error: 'Please enter at least one guest.' };
  if (guests > room.capacity) return { error: unit + ' sleeps a maximum of ' + room.capacity + '.' };

  let rate = room.rate;
  if (guests === 1 && String(room.singleAllowed).toLowerCase() === 'yes' && room.singleRate) {
    rate = room.singleRate;
  }

  return {
    firstName: firstName, lastName: lastName, email: email, phone: phone,
    unit: unit, unitType: room.type, guests: guests, notes: notes,
    checkin: checkin, checkout: checkout, checkinStr: checkinStr, checkoutStr: checkoutStr,
    nights: nights, rate: rate, total: rate * nights
  };
}

function lookupRoom(unit) {
  const rows = SpreadsheetApp.getActive().getSheetByName(TAB_ROOMS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === unit) {
      return {
        type: rows[i][1],
        capacity: Number(rows[i][2]),
        rate: Number(rows[i][3]),
        singleRate: rows[i][4] === '' ? 0 : Number(rows[i][4]),
        singleAllowed: rows[i][5],
        active: rows[i][6]
      };
    }
  }
  return null;
}

function boardUnitColumn(board, unit) {
  const header = board.getRange(1, 1, 1, board.getLastColumn()).getValues()[0];
  for (let c = 0; c < header.length; c++) {
    if (String(header[c]).trim() === unit) return c + 1;
  }
  return -1;
}

/** { 'yyyy-MM-dd': sheetRowNumber } for every date row on the board. */
function boardRowMap(board) {
  const count = board.getLastRow() - BOARD_HEADER_ROWS;
  const values = board.getRange(1 + BOARD_HEADER_ROWS, BOARD_DATE_COL, count, 1).getValues();
  const map = {};
  for (let i = 0; i < count; i++) {
    let v = values[i][0];
    if (!(v instanceof Date)) {
      const parsed = new Date(v);
      if (isNaN(parsed.getTime())) continue;
      v = parsed;
    }
    map[Utilities.formatDate(v, TZ, 'yyyy-MM-dd')] = i + 1 + BOARD_HEADER_ROWS;
  }
  return map;
}

/** { 'Tent 3': ['2026-10-01', ...], ... } for dates in [start, end] that are not Available. */
function blockedDates(startStr, endStr) {
  const board = SpreadsheetApp.getActive().getSheetByName(TAB_BOARD);
  const grid = board.getDataRange().getValues();
  const header = grid[0];
  const start = startStr ? parseISO(String(startStr).slice(0, 10)) : null;
  const end = endStr ? parseISO(String(endStr).slice(0, 10)) : null;
  const out = {};

  for (let r = BOARD_HEADER_ROWS; r < grid.length; r++) {
    let d = grid[r][BOARD_DATE_COL - 1];
    if (!(d instanceof Date)) {
      d = new Date(d);
      if (isNaN(d.getTime())) continue;
    }
    if (start && d < start) continue;
    if (end && d > end) continue;
    const key = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    for (let c = BOARD_FIRST_UNIT_COL - 1; c < header.length; c++) {
      if (String(grid[r][c]).trim() !== 'Available') {
        const unit = String(header[c]).trim();
        (out[unit] || (out[unit] = [])).push(key);
      }
    }
  }
  return out;
}

function dateKeys(checkin, nights) {
  const keys = [];
  for (let i = 0; i < nights; i++) {
    const d = new Date(checkin.getTime() + i * 86400000);
    keys.push(Utilities.formatDate(d, TZ, 'yyyy-MM-dd'));
  }
  return keys;
}

function parseISO(s) {
  return new Date(s + 'T00:00:00');
}

/** Next AMG-##### id, one above the highest already in the register. */
function nextBookingId(ss) {
  const reg = ss.getSheetByName(TAB_REGISTER);
  const last = reg.getLastRow();
  let max = 0;
  if (last > 1) {
    const ids = reg.getRange(2, 1, last - 1, 1).getValues();
    ids.forEach(function (r) {
      const m = String(r[0]).match(/^AMG-(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return 'AMG-' + String(max + 1).padStart(5, '0');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
