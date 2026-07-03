

// Google Apps Script Code
// Data is stored as JSON split across multiple cells in column A, because a
// single Google Sheets cell caps out at 50,000 characters. As the dashboard's
// data (tasks, notes, meetings, MOMs, materials grid, etc.) grows, one cell
// is no longer enough and writes were silently failing past that limit.
var CHUNK_SIZE = 45000;
var MAX_CHUNK_ROWS = 300;

function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = readChunks(sheet);

  // If the sheet is empty, return an empty object string
  if (!data) {
    data = "{}";
  }

  return ContentService.createTextOutput(data)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var jsonString = e.postData.contents;

    // Quick validation check
    JSON.parse(jsonString);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    writeChunks(sheet, jsonString);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function readChunks(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return "";
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  var parts = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v === "" || v === null || v === undefined) break;
    parts.push(String(v));
  }
  return parts.join("");
}

function writeChunks(sheet, jsonString) {
  var chunks = [];
  for (var i = 0; i < jsonString.length; i += CHUNK_SIZE) {
    chunks.push([jsonString.substring(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks = [[""]];

  // Clear a generous range first so leftover chunks from a longer previous
  // save don't linger past the end of the new (possibly shorter) data.
  sheet.getRange(1, 1, MAX_CHUNK_ROWS, 1).clearContent();
  sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
}
