/**
 * Work Time Tracker – data service for the GitHub Pages app.
 *
 * Sheet layout (first sheet, row 1 = headers):
 *   A  Timestamp (UTC)   ISO 8601 text, e.g. 2026-10-12T06:04:00.000Z  ← source of truth
 *   B  Action            "Clock In" / "Clock Out"
 *   C  Time zone         IANA zone of the device that recorded it, e.g. Europe/Helsinki
 *   D  Local time        Wall-clock time in that zone (real date value, handy for formulas)
 *
 * Times are logged rounded outward to the minute (clock-ins down, clock-outs up), and
 * a session that runs for 24 hours is stopped at the 24-hour mark the next time the
 * app talks to the server.
 *
 * One-time setup: run setupApiToken(), then migrateToUtc(), then roundAllTimes(), then
 * deploy as a web app (Execute as: Me, Who has access: Anyone).
 */

var HEADERS = ["Timestamp (UTC)", "Action", "Time zone", "Local time"];
var FORMATS = ["@", "@", "@", "yyyy-mm-dd hh:mm"];
var RECENT_COUNT = 6;
var MIN = 60 * 1000;
var CAP = 24 * 60 * MIN; // a session stops automatically after 24 hours

/* ---------------- HTTP endpoints ---------------- */

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "bad_request" });
  }

  var token = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  if (!token || req.token !== token) return json_({ ok: false, error: "unauthorized" });

  try {
    if (req.action === "state") return json_(withLock_(function() { autoStop_(getSheet_()); return state_(); }));
    if (req.action === "clock") return json_(clock_(req.type, req.tz));
    if (req.action === "edit")  return json_(edit_(req.row, req.at));
    return json_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return json_({ ok: false, error: "This endpoint only accepts requests from the tracker app." });
}

/* ---------------- Actions ---------------- */

function state_() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  var logs = [];

  if (last > 1) {
    var n = Math.min(RECENT_COUNT, last - 1);
    var start = last - n + 1;
    // Don't start the list halfway through a session: include the clock-in too.
    if (start > 2 && !isIn_(sheet.getRange(start, 2).getValue())) { start--; n++; }
    var values = sheet.getRange(start, 1, n, 3).getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      var d = toDate_(values[i][0]);
      logs.push({
        row: start + i,
        action: isIn_(values[i][1]) ? "in" : "out",
        at: d ? d.toISOString() : null,
        tz: String(values[i][2] || "")
      });
    }
  }

  return {
    ok: true,
    status: logs.length && logs[0].action === "in" ? "ON" : "OFF",
    logs: logs,
    serverTime: new Date().toISOString()
  };
}

function clock_(type, tz) {
  if (type !== "in" && type !== "out") throw new Error("Invalid clock type");
  tz = cleanTz_(tz);

  return withLock_(function() {
    var sheet = getSheet_();
    if (sheet.getLastRow() === 0) ensureHeaders_(sheet);
    // A session past 24 hours is closed first, so clocking in after one starts a new session.
    autoStop_(sheet);
    var last = sheet.getLastRow();
    var prev = last > 1 ? sheet.getRange(last, 1, 1, 2).getValues()[0] : null;
    var currentlyIn = !!prev && isIn_(prev[1]);

    // Only write if it actually changes the state, so double taps and retries can't create duplicates.
    if ((type === "in") !== currentlyIn) {
      var at = roundTime_(type === "in", new Date(), prev && isIn_(prev[1]), prev && toDate_(prev[0]));
      writeRow_(sheet, last + 1, at, type === "in", tz);
    }
    return state_();
  });
}

// Edits keep entries in the order they were recorded, which stops sessions from
// overlapping, and keep each session within 24 hours. The app checks the same
// rules first, with friendlier messages.
function edit_(row, at) {
  row = Number(row);
  var d = new Date(at);
  if (!(row > 1) || isNaN(d.getTime())) throw new Error("Invalid edit");

  return withLock_(function() {
    var sheet = getSheet_();
    var last = sheet.getLastRow();
    if (row > last) throw new Error("That entry no longer exists");
    var first = Math.max(2, row - 1), n = Math.min(last, row + 1) - first + 1;
    var rows = sheet.getRange(first, 1, n, 3).getValues();
    var cur = rows[row - first], prev = row > 2 ? rows[0] : null, next = row < last ? rows[rows.length - 1] : null;
    var isIn = isIn_(cur[1]);
    var t = roundTime_(isIn, d).getTime();
    var pt = prev && toDate_(prev[0]), nt = next && toDate_(next[0]);

    if (t > Date.now() + MIN) throw new Error("That time is in the future");
    if (pt && t < pt.getTime()) throw new Error("That's before the entry before it. Sessions can't overlap");
    if (nt && t > nt.getTime()) throw new Error("That's after the entry after it. Sessions can't overlap");
    if (!isIn && prev && isIn_(prev[1]) && pt && t - pt.getTime() > CAP) throw new Error("Sessions stop after 24 hours");
    if (isIn) {
      var end = !next ? Date.now() : !isIn_(next[1]) && nt ? nt.getTime() : null;
      if (end !== null && end - t > CAP) throw new Error("Sessions stop after 24 hours");
    }

    writeRow_(sheet, row, new Date(t), isIn, cleanTz_(cur[2]));
    return state_();
  });
}

// If the latest entry is a clock-in from 24 or more hours ago, clocks out at exactly
// 24 hours after it (after its rounded time, for entries logged before rounding).
function autoStop_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var vals = sheet.getRange(last, 1, 1, 3).getValues()[0];
  var start = toDate_(vals[0]);
  if (!isIn_(vals[1]) || !start) return;
  var end = roundTime_(true, start).getTime() + CAP;
  if (Date.now() < end) return;
  writeRow_(sheet, last + 1, new Date(end), false, cleanTz_(vals[2]));
}

/* ---------------- One-time setup (run from the editor) ---------------- */

/**
 * Rounds every logged time the way new ones are logged, and caps sessions at 24 hours.
 * Makes a backup copy of the sheet first. Safe to run more than once.
 */
function roundAllTimes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return;
  sheet.copyTo(ss).setName("Backup " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH.mm"));

  withLock_(function() {
    var range = sheet.getRange(2, 1, last - 1, 4);
    var values = range.getValues();
    var changed = 0, prevIn = null, prevAt = null;
    var out = values.map(function(r) {
      var d = toDate_(r[0]), isIn = isIn_(r[1]);
      if (!d) { prevAt = null; return r; }
      var at = roundTime_(isIn, d, prevIn, prevAt);
      prevIn = isIn; prevAt = at;
      if (at.getTime() === d.getTime() && typeof r[0] === "string") return r;
      changed++;
      var tz = cleanTz_(r[2]);
      return [at.toISOString(), r[1], tz, localString_(at, tz)];
    });
    range.setNumberFormats(out.map(function() { return FORMATS; }));
    range.setValues(out);
    autoStop_(sheet);
    Logger.log("Rounded " + changed + " of " + out.length + " rows.");
  });
}

/** Creates the access key the app asks for. Copy it from the execution log. */
function setupApiToken() {
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "").slice(0, 40);
  PropertiesService.getScriptProperties().setProperty("API_TOKEN", token);
  Logger.log("Your access key: " + token);
}

/**
 * Converts old rows (local times written in the script's time zone) to UTC.
 * Makes a backup copy of the sheet first. Safe to run more than once.
 * If your old entries were recorded in a different zone than the script's
 * (Project Settings → Time zone), put that zone in OLD_TZ, e.g. "Europe/Helsinki".
 */
function migrateToUtc() {
  var OLD_TZ = Session.getScriptTimeZone();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheet_();
  sheet.copyTo(ss).setName("Backup " + Utilities.formatDate(new Date(), OLD_TZ, "yyyy-MM-dd HH.mm"));
  ensureHeaders_(sheet);

  var last = sheet.getLastRow();
  if (last < 2) return;

  var ssTz = ss.getSpreadsheetTimeZone();
  var range = sheet.getRange(2, 1, last - 1, 4);
  var values = range.getValues();
  var skipped = 0;

  var out = values.map(function(r) {
    var a = r[0];
    if (typeof a === "string" && /Z$/.test(a.trim())) return r; // already UTC
    try {
      // Old code wrote local text in the script zone; Sheets then stored it as a date in the
      // spreadsheet zone. Recover the original wall-clock digits, then read them in OLD_TZ.
      var wall = (a instanceof Date)
        ? Utilities.formatDate(a, ssTz, "yyyy-MM-dd HH:mm:ss")
        : String(a).trim();
      if (wall.length === 16) wall += ":00";
      var d = Utilities.parseDate(wall, OLD_TZ, "yyyy-MM-dd HH:mm:ss");
      return [d.toISOString(), r[1], OLD_TZ, localString_(d, OLD_TZ)];
    } catch (err) {
      skipped++;
      return r;
    }
  });

  range.setNumberFormats(out.map(function() { return FORMATS; }));
  range.setValues(out);
  Logger.log("Migrated " + (out.length - skipped) + " rows, skipped " + skipped + ". Assumed zone: " + OLD_TZ);
}

/* ---------------- Helpers ---------------- */

// Clock-ins round down to the minute and clock-outs up, to cover the moment spent
// opening the app. Given the entry before it, a clock-in doesn't start before the
// clock-out it follows (rounding up can put that out a minute ahead), and a
// clock-out doesn't end more than 24 hours after its clock-in.
function roundTime_(isIn, date, prevIsIn, prevDate) {
  var t = date.getTime();
  var r = isIn ? Math.floor(t / MIN) * MIN : Math.ceil(t / MIN) * MIN;
  if (prevDate) {
    var p = prevDate.getTime();
    if (isIn && !prevIsIn && r < p && t > p - MIN) r = p;
    if (!isIn && prevIsIn && r - p > CAP) r = p + CAP;
  }
  return new Date(r);
}

function writeRow_(sheet, row, date, isIn, tz) {
  sheet.getRange(row, 1, 1, 4)
    .setNumberFormats([FORMATS])
    .setValues([[date.toISOString(), isIn ? "Clock In" : "Clock Out", tz, localString_(date, tz)]]);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function ensureHeaders_(sheet) {
  sheet.getRange(1, 1, 1, 4).setValues([HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function isIn_(action) {
  var a = String(action || "").toLowerCase();
  return a.indexOf("out") === -1 && a.indexOf("in") !== -1;
}

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v || "").trim();
  if (!s) return null;
  if (s.indexOf("T") !== -1) {
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  try { // unmigrated "yyyy-MM-dd HH:mm:ss" text
    return Utilities.parseDate(s.length === 16 ? s + ":00" : s, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  } catch (err) {
    return null;
  }
}

function cleanTz_(tz) {
  tz = String(tz || "");
  return /^[A-Za-z_]+(\/[A-Za-z0-9_+\-]+){0,2}$/.test(tz) && tz.length < 64 ? tz : Session.getScriptTimeZone();
}

function localString_(date, tz) {
  return Utilities.formatDate(date, tz, "yyyy-MM-dd HH:mm:ss");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
