/**
 * Lead webhook for the Assetz Miru & Miyo landing page.
 *
 * What it does: receives each form submission, emails it to you,
 * and appends a row to a Google Sheet (created automatically on first lead).
 *
 * NOTE: the page currently posts to the shared Sobha webhook, which already
 * works — deploying this script is OPTIONAL. Deploy it only if you want
 * Assetz leads in their own separate sheet with correct email subjects.
 *
 * HOW TO DEPLOY (5 minutes, one time):
 *  1. Go to https://script.google.com while logged in as kishanpvt26@gmail.com
 *  2. New project → delete the sample code → paste this whole file → Save (name it "Assetz Leads")
 *  3. Click Deploy → New deployment → gear icon → type: Web app
 *     - Execute as:        Me
 *     - Who has access:    Anyone
 *  4. Click Deploy, authorize when asked, and copy the Web app URL
 *     (looks like https://script.google.com/macros/s/AKfycb.../exec)
 *  5. Replace LEAD_ENDPOINT in index.html with that URL — done.
 */

var NOTIFY_EMAIL = "kishanpvt26@gmail.com";
var SHEET_NAME = "Assetz Miru Miyo Leads";
var PROJECT_LABEL = "Assetz Miru & Miyo";

function doPost(e) {
  var lead = {};
  try {
    lead = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ status: "error", message: "Invalid JSON" });
  }

  var rowError = "", mailError = "";
  try { appendToSheet(lead); } catch (err) { rowError = String(err); }
  try { sendEmail(lead); } catch (err) { mailError = String(err); }

  return jsonResponse({
    status: (rowError && mailError) ? "error" : "success",
    sheet: rowError || "ok",
    email: mailError || "ok"
  });
}

function sendEmail(lead) {
  var subject = "New Lead: " + (lead.name || "Unknown") + " — " + PROJECT_LABEL + " (" + (lead.formName || "Enquiry") + ")";
  var lines = [
    "Name:      " + (lead.name || "-"),
    "Phone:     " + (lead.countrycode || "") + " " + (lead.mobile || "-"),
    "Email:     " + (lead.email || "-"),
    "Form:      " + (lead.formName || "-"),
    "Source:    " + (lead.source || "-"),
    "",
    "Campaign:  " + (lead.utm_campaign || "-"),
    "UTM src:   " + (lead.utm_source || "-") + " / " + (lead.utm_medium || "-"),
    "Keyword:   " + (lead.utm_term || "-"),
    "gclid:     " + (lead.gclid || "-"),
    "",
    "Page:      " + (lead.page || "-"),
    "Time:      " + (lead.submitted_at || new Date().toISOString())
  ];
  MailApp.sendEmail(NOTIFY_EMAIL, subject, lines.join("\n"));
}

function appendToSheet(lead) {
  var files = DriveApp.getFilesByName(SHEET_NAME);
  var ss = files.hasNext()
    ? SpreadsheetApp.open(files.next())
    : SpreadsheetApp.create(SHEET_NAME);
  var sheet = ss.getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Time", "Name", "Phone", "Email", "Form", "Source",
                     "UTM Source", "UTM Medium", "UTM Campaign", "UTM Term", "gclid", "Page"]);
  }
  sheet.appendRow([
    lead.submitted_at || new Date().toISOString(),
    lead.name || "", (lead.countrycode || "") + " " + (lead.mobile || ""),
    lead.email || "", lead.formName || "", lead.source || "",
    lead.utm_source || "", lead.utm_medium || "", lead.utm_campaign || "",
    lead.utm_term || "", lead.gclid || "", lead.page || ""
  ]);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
