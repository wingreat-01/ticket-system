// ============================================================
//  IT TICKETING SYSTEM — Code.gs (Google Apps Script Backend)
// ============================================================

const SHEET_NAME = "Tickets";
const HEADERS = [
  "Ticket ID", "Date Submitted", "Requester Name", "Department",
  "Email", "Category", "Priority", "Subject", "Description",
  "Status", "Assigned To", "Resolution Notes", "Date Resolved", "Last Updated", "Signature"
];

// ── Entry Points ─────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("IT Ticketing System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var fn      = payload.fn;
    var params  = payload.params;
    var result;

    if      (fn === 'submitTicket')      result = submitTicket(params);
    else if (fn === 'getAllTickets')      result = getAllTickets();
    else if (fn === 'updateTicket')      result = updateTicket(params);
    else if (fn === 'deleteTicket')      result = deleteTicket(params);
    else if (fn === 'getDashboardStats') result = getDashboardStats();
    else result = JSON.stringify({ success: false, error: 'Unknown function: ' + fn });

    return ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helper: safe string conversion (handles Date objects) ────
function safeStr(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }
  return String(val);
}

// ── Helper: sheet row → plain string-only object ─────────────
function rowToObj(headers, row) {
  var obj = {};
  headers.forEach(function(h, i) { obj[h] = safeStr(row[i]); });
  return obj;
}

// ── Sheet Bootstrap ──────────────────────────────────────────
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);

    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground("#1a1a2e");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(8, 200);
    sheet.setColumnWidth(9, 300);
    sheet.setColumnWidth(12, 300);
  }

  return sheet;
}

// ── Ticket ID Generator ──────────────────────────────────────
function generateTicketId() {
  var sheet = getOrCreateSheet();
  var year  = new Date().getFullYear();
  var lastRow = sheet.getLastRow();

  // No tickets yet (only header row or empty)
  if (lastRow <= 1) return "TKT-" + year + "-0001";

  // Scan all existing Ticket IDs and find the highest number
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var max = 0;
  ids.forEach(function(row) {
    var id = String(row[0]);
    var match = id.match(/-(\d+)$/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > max) max = num;
    }
  });

  var pad = String(max + 1).padStart(4, "0");
  return "TKT-" + year + "-" + pad;
}

// ── Submit New Ticket ────────────────────────────────────────
// NOTE: Always return JSON.stringify — google.script.run cannot
//       serialize Date objects or complex nested objects reliably.
function submitTicket(dataJson) {
  try {
    var data = JSON.parse(dataJson);
    var sheet = getOrCreateSheet();
    var now = new Date();
    var tz = Session.getScriptTimeZone();
    var ticketId = generateTicketId();
    var nowStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");

    var row = [
      ticketId, nowStr,
      data.requesterName, data.department, data.email,
      data.category, data.priority, data.subject, data.description,
      "Open", "", "", "", nowStr
    ];

    sheet.appendRow(row);

    var newRowNum = sheet.getLastRow();
    var rowRange = sheet.getRange(newRowNum, 1, 1, HEADERS.length);
    var bg = data.priority === "Critical" ? "#ffd6d6"
           : data.priority === "High"     ? "#fff3cd"
           : data.priority === "Medium"   ? "#d6eaff"
           : "#f0fff0";
    rowRange.setBackground(bg);

    sendEmailNotification(data, ticketId);

    return JSON.stringify({ success: true, ticketId: ticketId });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ── Get All Tickets ──────────────────────────────────────────
// Converts every cell to a plain string (including Date cells)
// before JSON.stringify so nothing is lost in transit.
function getAllTickets() {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) return JSON.stringify({ success: true, tickets: [] });

    var headers = data[0];
    var tickets = data.slice(1).map(function(row) {
      return rowToObj(headers, row);
    }).reverse(); // newest first

    return JSON.stringify({ success: true, tickets: tickets });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ── Update Ticket ────────────────────────────────────────────
function updateTicket(payloadJson) {
  try {
    var payload  = JSON.parse(payloadJson);
    var ticketId = payload.ticketId;
    var updates  = payload.updates;
    var sheet    = getOrCreateSheet();
    var data     = sheet.getDataRange().getValues();
    var headers  = data[0];

    for (var i = 1; i < data.length; i++) {
      if (safeStr(data[i][0]) === ticketId) {
        var rowNum = i + 1;
        var fieldMap = {
          "Status":           headers.indexOf("Status") + 1,
          "Assigned To":      headers.indexOf("Assigned To") + 1,
          "Resolution Notes": headers.indexOf("Resolution Notes") + 1,
          "Priority":         headers.indexOf("Priority") + 1,
          "Signature":        headers.indexOf("Signature") + 1
        };

        Object.keys(updates).forEach(function(field) {
          if (fieldMap[field] !== undefined && fieldMap[field] > 0) {
            sheet.getRange(rowNum, fieldMap[field]).setValue(updates[field]);
          }
        });

        if (updates["Status"] === "Resolved" || updates["Status"] === "Closed") {
          var resolvedCol = headers.indexOf("Date Resolved") + 1;
          sheet.getRange(rowNum, resolvedCol).setValue(
            Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          );
        }

        var lastUpdatedCol = headers.indexOf("Last Updated") + 1;
        sheet.getRange(rowNum, lastUpdatedCol).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
        );

        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: "Ticket not found" });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ── Delete Ticket ────────────────────────────────────────────
function deleteTicket(ticketId) {
  try {
    var sheet = getOrCreateSheet();
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (safeStr(data[i][0]) === ticketId) {
        sheet.deleteRow(i + 1);
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: "Ticket not found" });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ── Dashboard Stats ──────────────────────────────────────────
function getDashboardStats() {
  try {
    var sheet = getOrCreateSheet();
    var data  = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return JSON.stringify({
        success: true,
        stats: { total: 0, open: 0, inProgress: 0, resolved: 0, closed: 0, critical: 0 },
        categories: {}, departments: {}
      });
    }

    var headers     = data[0];
    var statusIdx   = headers.indexOf("Status");
    var priorityIdx = headers.indexOf("Priority");
    var categoryIdx = headers.indexOf("Category");
    var deptIdx     = headers.indexOf("Department");

    var stats       = { total: 0, open: 0, inProgress: 0, resolved: 0, closed: 0, critical: 0 };
    var categories  = {};
    var departments = {};

    data.slice(1).forEach(function(row) {
      stats.total++;
      var status   = safeStr(row[statusIdx]);
      var priority = safeStr(row[priorityIdx]);
      var category = safeStr(row[categoryIdx]);
      var dept     = safeStr(row[deptIdx]);

      if (status === "Open")             stats.open++;
      else if (status === "In Progress") stats.inProgress++;
      else if (status === "Resolved")    stats.resolved++;
      else if (status === "Closed")      stats.closed++;

      if (priority === "Critical") stats.critical++;
      if (category) categories[category]   = (categories[category]   || 0) + 1;
      if (dept)     departments[dept]       = (departments[dept]       || 0) + 1;
    });

    return JSON.stringify({ success: true, stats: stats, categories: categories, departments: departments });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ── Email Notification (Optional) ───────────────────────────
function sendEmailNotification(data, ticketId) {
  try {
    if (!data.email) return;
    var subject = "[IT Ticket " + ticketId + "] " + data.subject;
    var body = "Hello " + data.requesterName + ",\n\n"
      + "Your IT support ticket has been submitted successfully.\n\n"
      + "Ticket ID   : " + ticketId + "\n"
      + "Category    : " + data.category + "\n"
      + "Priority    : " + data.priority + "\n"
      + "Subject     : " + data.subject + "\n"
      + "Description : " + data.description + "\n"
      + "Status      : Open\n\n"
      + "Our IT team will review your request and get back to you shortly.\n\n"
      + "Thank you,\nIT Support Team";
    MailApp.sendEmail({ to: data.email, subject: subject, body: body });
  } catch (e) {
    Logger.log("Email error: " + e.message);
  }
}
