// Execution Capacity Scorecard — submission handler.
// Paste this into a Google Sheet's Extensions > Apps Script, then deploy as a Web App.
// See README-notifications.md (or the setup instructions from Claude) for exact steps.

const NOTIFY_EMAIL = "kristina@neurodivergentventures.com"; // where "new submission" pings go

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  logToSheet(data);
  notifyConsultant(data);
  emailClientReport(data);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function logToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Submissions") || createSheet(ss);
  sheet.appendRow([
    new Date(data.submittedAt || Date.now()),
    data.client.name,
    data.client.email,
    data.client.consultant || "",
    data.rec.level,
    data.rec.program,
    (data.top3 || []).map(d => d.name).join(", "),
  ]);
}

function createSheet(ss) {
  const sheet = ss.insertSheet("Submissions");
  sheet.appendRow(["Date", "Name", "Email", "Consultant", "Friction Level", "Program", "Top Focus Areas"]);
  return sheet;
}

function notifyConsultant(data) {
  const subject = `New scorecard submitted: ${data.client.name}`;
  const body = [
    `${data.client.name} just completed the Execution Capacity Scorecard.`,
    "",
    `Email: ${data.client.email}`,
    `Referring consultant: ${data.client.consultant || "Not specified"}`,
    `Execution Friction Level: ${data.rec.level} (${data.rec.program})`,
    `Top Focus Areas: ${(data.top3 || []).map(d => d.name).join(", ") || "None"}`,
  ].join("\n");
  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

function emailClientReport(data) {
  const c = data.client;
  const lines = [
    `Hi ${(c.name || "").split(" ")[0] || "there"},`,
    "",
    "Here's a summary of your Execution Capacity Scorecard results.",
    "",
    `Execution Friction Level: ${data.rec.level} — ${data.rec.program}`,
    "",
    "Top Focus Areas:",
    ...(data.top3 || []).map((d, i) => `${i + 1}. ${d.name} — ${d.action}`),
    "",
    (data.ninetyDay && data.ninetyDay.length)
      ? ["Your 90-Day Focus:", ...data.ninetyDay.map((item, i) => `${i + 1}. ${item}`)].join("\n")
      : "Every area came back a strength — the focus for the next 90 days is maintaining what's working.",
    "",
    "We'll be in touch to talk through what this means and how we tackle it together.",
  ];
  MailApp.sendEmail(c.email, "Your Execution Capacity Scorecard Results", lines.join("\n"));
}
