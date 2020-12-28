function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3').activate();
  spreadsheet.getCurrentCell().setValue('TRUE');
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A4').activate();
  spreadsheet.getCurrentCell().setValue('TRUE');
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A5').activate();
  spreadsheet.getCurrentCell().setValue('TRUE');
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A6').activate();
  spreadsheet.getCurrentCell().setValue('TRUE');
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A7').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A2').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function UntitledMacro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:A7').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function m7() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7').activate();
  spreadsheet.getRange('A7').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .setHelpText('Enter Checked or Unchecked.')
  .requireCheckbox('Checked', 'Unchecked')
  .build());
};

function test_clear_formatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getActiveRange().getDataRegion().activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clearFormat();
};