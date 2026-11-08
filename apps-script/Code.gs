// Google Apps Script bound to a Google Sheet. Deploy as Web app (Execute as: Me, Access: Anyone).
// Columns: A Timestamp | B Group | C Poster Leader | D Email | E Topic
const DEADLINE = new Date("2026-11-13T23:59:59+01:00");
const HEADER = ["Timestamp", "Group", "Poster Leader", "Email", "Topic"];

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  return sh;
}

function taken_(sh) {
  const taken = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[4]) taken[r[4]] = r[1]; });
  return taken;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ taken: taken_(sheet_()), closed: new Date() > DEADLINE });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (new Date() > DEADLINE) return json_({ ok: false, error: "Topic selection is closed." });
    const sh = sheet_(), taken = taken_(sh);
    if (taken[d.topic]) return json_({ ok: false, error: `"${d.topic}" was just taken by ${taken[d.topic]}. Please choose another topic.` });
    if (Object.values(taken).includes(d.group)) return json_({ ok: false, error: `${d.group} has already chosen a topic.` });
    sh.appendRow([new Date(), d.group, d.name, d.email, d.topic]);
  } finally { lock.releaseLock(); }
  MailApp.sendEmail(d.email, `PM Poster Session – ${d.group}: topic confirmed`,
    `Dear ${d.name},\n\n${d.group} will present the topic "${d.topic}" at the Poster Session ` +
    `on Friday 11.12.2026, 15:45–18:15.\n\nPlease share this with your group.\n\nBest regards,\nIntroduction to Project Management Teaching Team`);
  return json_({ ok: true });
}
