// Paste this whole file into Extensions > Apps Script for the "One Child Interviews" sheet.

// Bookings allowed per slot, keyed by the date label exactly as it appears in
// the "<date label> at <time>" timeSlot string. Must stay in sync with the
// `capacity` field per date in the `availability` array in script.js.
var DATE_CAPACITY = {
  "Tuesday, July 28, 2026": 1,
  "Thursday, July 30, 2026": 1,
  "Friday, July 31, 2026": 5,
};
var DEFAULT_CAPACITY = 1;

function getCapacityForSlot(timeSlot) {
  var dateLabel = String(timeSlot).split(" at ")[0];
  return DATE_CAPACITY.hasOwnProperty(dateLabel) ? DATE_CAPACITY[dateLabel] : DEFAULT_CAPACITY;
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var bookedSlots = getBookedSlots(sheet);

  return ContentService
    .createTextOutput(JSON.stringify({ bookedSlots: bookedSlots }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "The system is busy, please try again." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    var existingCount = getBookedSlots(sheet).filter(function (slot) {
      return slot === data.timeSlot;
    }).length;
    var capacity = getCapacityForSlot(data.timeSlot);

    if (existingCount >= capacity) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: "error",
          message: "Sorry, that time slot was just booked by someone else. Please choose another.",
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow([
      data.name,
      data.email,
      data.phone,
      data.timeSlot,
      new Date(),
    ]);

    sortSheetByInterviewTime(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getBookedSlots(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Column D = Time Slot
  var values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  return values.map(function (row) { return row[0]; }).filter(String);
}

// Sorts data rows (everything below the header) by interview date/time,
// derived from the "<date label> at <start time> - <end time>" timeSlot
// string in column D. Reorders whole rows in place so every other column
// stays aligned with its row - this only changes display order and has no
// effect on the exact-string booking counts used by the capacity checks.
function sortSheetByInterviewTime(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // fewer than 2 data rows - nothing to reorder

  var numCols = sheet.getLastColumn();
  var range = sheet.getRange(2, 1, lastRow - 1, numCols);
  var values = range.getValues();

  values.sort(function (a, b) {
    return getSlotSortKey(a[3]) - getSlotSortKey(b[3]);
  });

  range.setValues(values);
}

function getSlotSortKey(timeSlot) {
  var parts = String(timeSlot).split(" at ");
  var dateLabel = parts[0] || "";
  var startTime = (parts[1] || "").split(" - ")[0];

  // Strip the leading weekday name (e.g. "Monday, ") so Date parsing is reliable.
  var dateOnly = dateLabel.replace(/^[A-Za-z]+,\s*/, "");
  var parsed = new Date(dateOnly + " " + startTime);

  // Unparseable slots sort last instead of throwing, so one bad row can't
  // break sorting for everyone else.
  return isNaN(parsed.getTime()) ? Infinity : parsed.getTime();
}
