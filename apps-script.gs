// Paste this whole file into Extensions > Apps Script for the "One Child Interviews" sheet.

// Bookings allowed per slot, keyed by the date label exactly as it appears in
// the "<date label> at <time>" timeSlot string. Must stay in sync with the
// `capacity` field per date in the `availability` array in script.js.
var DATE_CAPACITY = {
  "Thursday, August 13, 2026": 3,
  "Friday, August 14, 2026": 5,
  "Monday, August 17, 2026": 5,
  "Thursday, August 20, 2026": 4,
  "Friday, August 21, 2026": 4,
  "Monday, August 24, 2026": 8,
  "Tuesday, August 25, 2026": 8,
};
var DEFAULT_CAPACITY = 1;

// Short-lived cache of the booked-slots list so rapid date-switching on the
// frontend doesn't trigger a fresh sheet read on every click. doPost always
// reads the sheet directly (never the cache) so double-booking prevention is
// unaffected, and clears this cache on a successful booking so the next
// doGet reflects it immediately instead of waiting out the full TTL.
var BOOKED_SLOTS_CACHE_KEY = "bookedSlots";
var BOOKED_SLOTS_CACHE_TTL_SECONDS = 20;

function getCapacityForSlot(timeSlot) {
  var dateLabel = String(timeSlot).split(" at ")[0];
  return DATE_CAPACITY.hasOwnProperty(dateLabel) ? DATE_CAPACITY[dateLabel] : DEFAULT_CAPACITY;
}

// Simple trigger: fires automatically whenever a cell is edited directly in
// this spreadsheet (not for edits made by the script itself, e.g. by
// syncRecruitmentApplicants() writing into Recruitment_applicants - simple
// triggers don't re-fire for those). Re-runs syncRecruitmentApplicants()
// whenever an edit lands in column F (Interview Status) or column G
// (Assigned Interviewer) on the main sheet, so Stage and Interviewer
// changes reflect in Recruitment_applicants within a second or two, not
// just on new form submissions.
//
// Runs in Apps Script's restricted "simple trigger" mode, which only
// allows access to the SAME spreadsheet the edit happened in - exactly
// what syncRecruitmentApplicants() needs, so no new authorization prompt
// should be required.
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  if (sheet.getSheetId() !== mainSheet.getSheetId()) return;
  if (range.getRow() < 2) return; // ignore header row edits

  var editedCol = range.getColumn();
  var lastEditedCol = editedCol + range.getNumColumns() - 1;

  var touchesColF = editedCol <= 6 && lastEditedCol >= 6; // Interview Status
  var touchesColG = editedCol <= 7 && lastEditedCol >= 7; // Assigned Interviewer

  if (!touchesColF && !touchesColG) return;

  syncRecruitmentApplicants();
}

function doGet(e) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(BOOKED_SLOTS_CACHE_KEY);
  var bookedSlots;

  if (cached) {
    bookedSlots = JSON.parse(cached);
  } else {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    bookedSlots = getBookedSlots(sheet);
    cache.put(BOOKED_SLOTS_CACHE_KEY, JSON.stringify(bookedSlots), BOOKED_SLOTS_CACHE_TTL_SECONDS);
  }

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

    if (hasDuplicateSubmission(sheet, data.name, data.phone, data.timeSlot)) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: "error",
          message: "You've already submitted for this time slot.",
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
    CacheService.getScriptCache().remove(BOOKED_SLOTS_CACHE_KEY);
    syncRecruitmentApplicants();

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

// Rejects a resubmission from the same person (by Name + Phone) for the
// same exact Time Slot. Name is compared case-insensitively; phone is
// compared with formatting characters stripped, so "(555) 123-4567" and
// "5551234567" are treated as the same number.
function hasDuplicateSubmission(sheet, name, phone, timeSlot) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  var normalizedName = String(name).trim().toLowerCase();
  var normalizedPhone = String(phone).replace(/\D/g, "");

  var rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A:D = Name, Email, Phone, Time Slot
  return rows.some(function (row) {
    var rowName = String(row[0]).trim().toLowerCase();
    var rowPhone = String(row[2]).replace(/\D/g, "");
    var rowTimeSlot = row[3];
    return rowName === normalizedName && rowPhone === normalizedPhone && rowTimeSlot === timeSlot;
  });
}

// Syncs interviewee rows from the main sheet into Recruitment_applicants,
// matched by Email (case-insensitive). New emails get a new row with
// Applicant/Email/Interviewer/Date of Interview filled in. Existing rows
// only have those same three fields refreshed - French Speaking (Fluent),
// Phase, Stage, and Notes are never touched, even when blank, since staff
// fill those in manually and a blank may be intentional.
//
// Safe to run repeatedly (idempotent by email) - this same function is used
// both for the one-time catch-up and automatically after every submission.
// Does not call SpreadsheetApp.getUi(), since it also runs from doPost's
// headless context, where UI calls would throw.
function syncRecruitmentApplicants() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheets()[0];
  var appTab = ss.getSheetByName("Recruitment_applicants");
  if (!appTab) return;

  var mainLastRow = mainSheet.getLastRow();
  if (mainLastRow < 2) return;

  var mainRows = mainSheet.getRange(2, 1, mainLastRow - 1, 7).getValues(); // A:G

  var appLastRow = appTab.getLastRow();
  var emailToRow = {}; // lowercased email -> row number in appTab
  if (appLastRow >= 2) {
    var appEmails = appTab.getRange(2, 2, appLastRow - 1, 1).getValues(); // column B = Email
    for (var i = 0; i < appEmails.length; i++) {
      var existingEmail = String(appEmails[i][0]).trim().toLowerCase();
      if (existingEmail) emailToRow[existingEmail] = i + 2;
    }
  }

  var addedCount = 0;
  var updatedCount = 0;

  mainRows.forEach(function (row) {
    var name = row[0];
    var email = String(row[1] || "").trim();
    var timeSlot = row[3];
    var interviewer = row[6]; // Column G - Assigned Interviewer

    if (!email) return;

    var key = email.toLowerCase();
    var existingRow = emailToRow[key];

    if (existingRow) {
      appTab.getRange(existingRow, 1).setValue(name);       // A - Applicant
      appTab.getRange(existingRow, 5).setValue(interviewer); // E - Interviewer
      appTab.getRange(existingRow, 6).setValue(timeSlot);    // F - Date of Interview
      updatedCount++;
    } else {
      appTab.appendRow([name, email, "", "", interviewer, timeSlot, "", ""]);
      emailToRow[key] = appTab.getLastRow();
      addedCount++;
    }
  });

  Logger.log("Recruitment_applicants sync: added " + addedCount + ", updated " + updatedCount + ".");
}

// One-time setup: adds dropdown data validation to the main sheet's column F
// (Interview Status), covering rows 2-1000 so newly added rows are
// automatically covered without re-running this. Options mirror what was
// already configured manually on this column, plus "Interview Scheduled"
// (previously missing here, though already present in the
// Recruitment_applicants "Stage" dropdown below) so that value is valid on
// both tabs.
function setupMainSheetInterviewStatusDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheets()[0];

  var STATUS_OPTIONS = [
    "Interview Scheduled",
    "Interviewed - Pending",
    "Interviewed - Accepted",
    "Interviewed - Rejected",
    "NO SHOW",
    "TBD - RESCHEDULED",
  ];

  var ROW_START = 2;
  var ROW_COUNT = 999; // rows 2 through 1000

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  mainSheet.getRange(ROW_START, 6, ROW_COUNT, 1).setDataValidation(statusRule);

  Logger.log("Dropdown validation applied to main sheet column F (Interview Status), rows 2-1000.");
}

// One-time setup: adds dropdown data validation to Recruitment_applicants
// columns C (French Speaking), D (Phase), E (Interviewer), and G (Stage),
// covering rows 2-1000 in a single rule per column so newly added rows
// (via sync or manual entry) are automatically covered without re-running
// this. Interviewer options are pulled live from the Interviewers tab
// rather than hardcoded, so adding/removing an interviewer there updates
// this dropdown automatically - matching the same approach used for the
// Assigned Interviewer dropdown on the main sheet.
function setupRecruitmentApplicantsDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var appTab = ss.getSheetByName("Recruitment_applicants");
  if (!appTab) {
    throw new Error("Recruitment_applicants tab not found.");
  }

  var interviewersTab = ss.getSheetByName("Interviewers");
  if (!interviewersTab) {
    throw new Error("Interviewers tab not found.");
  }

  var FRENCH_SPEAKING_OPTIONS = ["YES", "NO", "N/A", "Basic"];
  var PHASE_OPTIONS = ["Phase 1", "Phase 2"];
  var STAGE_OPTIONS = [
    "Interview Scheduled",
    "Interviewed – Pending",
    "Interviewed – Accepted",
    "Interviewed – Rejected",
    "NO SHOW",
    "TBD - RESCHEDULED",
  ];

  var ROW_START = 2;
  var ROW_COUNT = 999; // rows 2 through 1000

  // Column C - French Speaking (Fluent)
  var frenchRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FRENCH_SPEAKING_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  appTab.getRange(ROW_START, 3, ROW_COUNT, 1).setDataValidation(frenchRule);

  // Column D - Phase
  var phaseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PHASE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  appTab.getRange(ROW_START, 4, ROW_COUNT, 1).setDataValidation(phaseRule);

  // Column E - Interviewer (pulled live from Interviewers!A2:A, not hardcoded)
  var interviewersLastRow = interviewersTab.getLastRow();
  var interviewerNameRange = interviewersTab.getRange(2, 1, Math.max(interviewersLastRow - 1, 1), 1);
  var interviewerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(interviewerNameRange, true)
    .setAllowInvalid(false)
    .build();
  appTab.getRange(ROW_START, 5, ROW_COUNT, 1).setDataValidation(interviewerRule);

  // Column G - Stage
  var stageRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STAGE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  appTab.getRange(ROW_START, 7, ROW_COUNT, 1).setDataValidation(stageRule);

  Logger.log("Dropdown validation rules applied to Recruitment_applicants (rows 2-1000): columns C, D, E, G.");
}

// One-time backfill: pulls French Speaking, Phase, Stage, and Notes from an
// older separate spreadsheet into this sheet's Recruitment_applicants tab,
// matched by Email (case-insensitive). Only fills cells that are currently
// BLANK here - never overwrites existing data. Does not touch Applicant,
// Email, Interviewer, or Date of Interview, since those are already kept
// current by syncRecruitmentApplicants().
var OLD_SHEET_ID = "119ETHN0YGMdWo4bQWj1ekyhRSWqAMevCFrToAo4sVKM";

function backfillFromOldSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var appTab = ss.getSheetByName("Recruitment_applicants");
  if (!appTab) {
    throw new Error("Recruitment_applicants tab not found.");
  }

  var oldSs = SpreadsheetApp.openById(OLD_SHEET_ID);
  var oldSheet = oldSs.getSheets()[0];
  var oldLastRow = oldSheet.getLastRow();
  if (oldLastRow < 2) {
    Logger.log("Old sheet has no data rows; nothing to backfill.");
    return;
  }

  var oldRows = oldSheet.getRange(2, 1, oldLastRow - 1, 8).getValues(); // A:H

  // Build email -> best row map. Rows are processed in sheet order (top to
  // bottom = oldest to newest, since new attempts get appended at the
  // bottom). A row with a non-blank Stage always wins over the current
  // pick, so the LAST non-blank-Stage row for that email ends up stored.
  // If no row for that email ever has a Stage, the LAST row overall is
  // kept as a fallback, for the same "later = more current" reasoning.
  var emailToRow = {};
  oldRows.forEach(function (row) {
    var email = String(row[1] || "").trim().toLowerCase();
    if (!email) return;

    var existing = emailToRow[email];
    if (!existing) {
      emailToRow[email] = row;
      return;
    }

    var newStageBlank = !String(row[6]).trim();
    var existingStageBlank = !String(existing[6]).trim();

    if (!newStageBlank) {
      // This row has a Stage - always take it (last non-blank Stage wins).
      emailToRow[email] = row;
    } else if (existingStageBlank) {
      // Neither has a Stage yet - keep tracking the latest row as fallback.
      emailToRow[email] = row;
    }
    // else: existing already has a Stage and this row doesn't - keep existing.
  });

  var appLastRow = appTab.getLastRow();
  if (appLastRow < 2) {
    Logger.log("Recruitment_applicants has no data rows; nothing to backfill.");
    return;
  }

  var appRows = appTab.getRange(2, 1, appLastRow - 1, 8).getValues(); // A:H
  var matchedCount = 0;
  var updatedRowCount = 0;
  var updatedCellCount = 0;

  for (var i = 0; i < appRows.length; i++) {
    var row = appRows[i];
    var email = String(row[1] || "").trim().toLowerCase();
    if (!email) continue;

    var oldRow = emailToRow[email];
    if (!oldRow) continue;

    matchedCount++;
    var sheetRow = i + 2;
    var rowChanged = false;

    // Column C - French Speaking (Fluent)
    if (!String(row[2]).trim() && String(oldRow[2]).trim()) {
      appTab.getRange(sheetRow, 3).setValue(oldRow[2]);
      updatedCellCount++;
      rowChanged = true;
    }
    // Column D - Phase
    if (!String(row[3]).trim() && String(oldRow[3]).trim()) {
      appTab.getRange(sheetRow, 4).setValue(oldRow[3]);
      updatedCellCount++;
      rowChanged = true;
    }
    // Column G - Stage
    if (!String(row[6]).trim() && String(oldRow[6]).trim()) {
      appTab.getRange(sheetRow, 7).setValue(oldRow[6]);
      updatedCellCount++;
      rowChanged = true;
    }
    // Column H - Notes
    if (!String(row[7]).trim() && String(oldRow[7]).trim()) {
      appTab.getRange(sheetRow, 8).setValue(oldRow[7]);
      updatedCellCount++;
      rowChanged = true;
    }

    if (rowChanged) updatedRowCount++;
  }

  Logger.log(
    "Backfill from old sheet: " + matchedCount + " applicant(s) matched by email, " +
    updatedRowCount + " row(s) updated, " + updatedCellCount + " cell(s) filled."
  );
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
    "IMPORTANT: Your camera must be ON for the duration of the interview.\n\n" +
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
