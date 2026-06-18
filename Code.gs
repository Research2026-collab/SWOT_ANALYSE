// SWOT Workshop Board — Google Apps Script Backend
// Daten werden in das aktive Google Sheet geschrieben.
// Sheet-Struktur: Spalten A=Timestamp, B=Quadrant, C=Text

var SHEET_NAME = 'SWOT-Daten';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SWOT-Workshop')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getEntries() {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var entries = { S: [], W: [], O: [], T: [] };
  for (var i = 1; i < data.length; i++) {
    var q = data[i][1];
    var text = data[i][2];
    if (q && text && entries[q] !== undefined) {
      entries[q].push(text);
    }
  }
  return entries;
}

function addEntry(quadrant, text) {
  if (!text || text.trim() === '') return false;
  if (['S', 'W', 'O', 'T'].indexOf(quadrant) === -1) return false;
  var sheet = getOrCreateSheet();
  sheet.appendRow([new Date(), quadrant, text.trim()]);
  return true;
}

function clearAll() {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return true;
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Quadrant', 'Text']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sheet;
}
