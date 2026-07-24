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

// Triggered by the "Send Confirmation" drawing/button on the main sheet.
// Sends the selected interviewee a confirmation email using their assigned
// interviewer's Zoom link (Interviewers tab), then stamps column H
// (Confirmation Sent). Column F (Interview Status) is unrelated and untouched.
function sendSelectedConfirmation() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    ui.alert("Select an interviewee row first (not the header row).");
    return;
  }

  var rowData = sheet.getRange(row, 1, 1, 8).getValues()[0];
  var name = rowData[0];
  var email = rowData[1];
  var timeSlot = rowData[3];
  var interviewerName = rowData[6]; // Column G - Assigned Interviewer
  var confirmationSent = rowData[7]; // Column H - Confirmation Sent

  if (!name || !email) {
    ui.alert("This row is missing a name or email.");
    return;
  }

  if (confirmationSent) {
    var proceed = ui.alert(
      "Confirmation already sent on " + confirmationSent + ". Send again?",
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  }

  if (!interviewerName) {
    ui.alert("No interviewer assigned in column G for this row.");
    return;
  }

  var zoomLink = getZoomLinkForInterviewer(ss, interviewerName);
  if (!zoomLink) {
    ui.alert("No matching Zoom link found for interviewer \"" + interviewerName + "\" on the Interviewers tab.");
    return;
  }

  var subject = "Your Interview Confirmation - " + timeSlot;
  var body =
    "Hi " + name + ",\n\n" +
    "This confirms your interview is scheduled for " + timeSlot + " (Eastern Time).\n\n" +
    "Your interviewer is " + interviewerName + ". Join using this Zoom link:\n" + zoomLink + "\n\n" +
    "See you then!";

  MailApp.sendEmail(email, subject, body);

  sheet.getRange(row, 8).setValue(new Date());
  ui.alert("Confirmation email sent to " + email + ".");
}

// Looks up the Zoom link for an interviewer by exact name match against
// column A of the Interviewers tab (column B holds the Zoom link).
function getZoomLinkForInterviewer(ss, interviewerName) {
  var tab = ss.getSheetByName("Interviewers");
  if (!tab) return null;

  var lastRow = tab.getLastRow();
  if (lastRow < 2) return null;

  var data = tab.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(interviewerName).trim()) {
      return data[i][1];
    }
  }
  return null;
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
