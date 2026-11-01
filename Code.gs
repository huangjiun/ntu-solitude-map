const SHEET_NAME = "投稿";
const ADMIN_SHEET_NAME = "Admins";
const GOOGLE_CLIENT_ID = "替換成你的 Google OAuth Web Client ID";

function doGet(e) {
  const action = e.parameter.action || "approved";
  if (action !== "approved" && !isAdmin_(e.parameter.credential)) return json_({ error: "unauthorized" });
  const rows = getRows_().filter(r => action === "approved" ? r.status === "approved" : r.status === "pending");
  return json_({ items: rows });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents || "{}");
  if (data.action === "submit") {
    sheet_().appendRow([Utilities.getUuid(), new Date(), data.name, data.lng, data.lat, data.reason || "", "pending"]);
    return json_({ ok: true });
  }
  if (data.action === "pending" && isAdmin_(data.credential)) {
    return json_({ items: getRows_().filter(r => r.status === "pending") });
  }
  if (data.action === "review" && isAdmin_(data.credential) && ["approved", "rejected"].includes(data.status)) {
    const sheet = sheet_(), values = sheet.getDataRange().getValues();
    const row = values.findIndex((r, i) => i > 0 && r[0] === data.id);
    if (row > 0) sheet.getRange(row + 1, 7).setValue(data.status);
    return json_({ ok: row > 0 });
  }
  return json_({ error: "unauthorized" });
}

function isAdmin_(credential) {
  if (!credential) return false;
  try {
    const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(credential), { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return false;
    const token = JSON.parse(response.getContentText());
    if (token.aud !== GOOGLE_CLIENT_ID || token.email_verified !== "true") return false;
    return admins_().includes(String(token.email).toLowerCase());
  } catch (error) {
    return false;
  }
}

function sheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActive().insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(["id", "createdAt", "name", "lng", "lat", "reason", "status"]);
  return sheet;
}
function admins_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ADMIN_SHEET_NAME) || SpreadsheetApp.getActive().insertSheet(ADMIN_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(["email"]);
  return sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat().filter(Boolean).map(e => String(e).trim().toLowerCase());
}
function getRows_() {
  const values = sheet_().getDataRange().getValues();
  return values.slice(1).filter(r => r[0]).map(r => ({ id: r[0], createdAt: r[1], name: r[2], lng: Number(r[3]), lat: Number(r[4]), reason: r[5], status: r[6] }));
}
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
