function TAGS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('G2:G101').activate();
  spreadsheet.getRange('G2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
};

function TAGSrelativos() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('G2:G101').activate();
  spreadsheet.getRange('G2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
};