const SHEET_NAME = "投稿";
const ADMIN_EMAILS = ["請替換成你的 Google 帳號 email"];

function doGet(e) {
  const action = e.parameter.action || "approved";
  if (action !== "approved" && !isAdmin_()) return json_({ error: "unauthorized" });
  const rows = getRows_().filter(r => action === "approved" ? r.status === "approved" : r.status === "pending");
  return json_({ items: rows });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents || "{}");
  if (data.action === "submit") {
    const sheet = sheet_();
    sheet.appendRow([Utilities.getUuid(), new Date(), data.name, data.lng, data.lat, data.reason || "", "pending"]);
    return json_({ ok: true });
  }
  if (data.action === "review" && isAdmin_()) {
    const sheet = sheet_(), values = sheet.getDataRange().getValues();
    const row = values.findIndex((r, i) => i > 0 && r[0] === data.id);
    if (row > 0) sheet.getRange(row + 1, 7).setValue(data.status);
    return json_({ ok: row > 0 });
  }
  return json_({ error: "unauthorized" });
}

function sheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActive().insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(["id", "createdAt", "name", "lng", "lat", "reason", "status"]);
  return sheet;
}
function getRows_() {
  const values = sheet_().getDataRange().getValues();
  return values.slice(1).filter(r => r[0]).map(r => ({ id: r[0], createdAt: r[1], name: r[2], lng: Number(r[3]), lat: Number(r[4]), reason: r[5], status: r[6] }));
}
function isAdmin_() { return ADMIN_EMAILS.includes(Session.getActiveUser().getEmail()); }
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
