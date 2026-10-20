/**
 * Guild Screening Test: receives submissions from the GitHub page and saves them in this Google Sheet.
 *
 * 1. Change SECRET below to a long word only you know.
 * 2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 3. Paste the web app URL into SUBMIT_URL at the top of the script in index.html.
 * 4. To collect submissions for your results page, open:  <web app URL>?key=<your SECRET>
 */
const SECRET = "change-this-to-a-long-secret-word";
const SHEET_NAME = "Submissions";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Received", "Submission ID", "Name", "Section", "Minutes", "Data (JSON)"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Stops a name like "=HYPERLINK(...)" from being treated as a formula.
function safe_(v, max) {
  const s = String(v == null ? "" : v).slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const rec = JSON.parse(e.postData.contents);
    if (!rec || typeof rec !== "object" || !rec.answers) return ContentService.createTextOutput("ignored");
    let json = JSON.stringify(rec);
    if (json.length > 49000) json = json.slice(0, 49000); // a sheet cell holds 50,000 characters
    sheet_().appendRow([new Date(), safe_(rec.id, 64), safe_(rec.name, 120), safe_(rec.section, 120), Number(rec.minutes) || 0, safe_(json, 49000)]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== SECRET) {
    return ContentService.createTextOutput("Add ?key=YOUR_SECRET to the end of this address to see the submissions.");
  }
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = [];
  rows.forEach(function (r) {
    try { out.push(JSON.parse(String(r[5]).replace(/^'/, ""))); } catch (err) { /* skip a damaged row */ }
  });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
