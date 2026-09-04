const SHEET_NAME = "Responses";

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(["Timestamp","Name","Response"]);
}

function doPost(e) {
  setup();
  const raw = e.postData && e.postData.contents ? e.postData.contents : "{}";
  const data = JSON.parse(raw);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    .appendRow([new Date(), data.name || "Oyin", data.response || ""]);
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  setup();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sh.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[2]).map(r => ({
    time: new Date(r[0]).toISOString(),
    name: r[1],
    response: r[2]
  }));
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}