function denyAndMoveRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var buttonSheet = ss.getSheetByName('Leave Request Approval'); 
  
  if (!buttonSheet) {
    Logger.log('ButtonSheet not found.');
    return;
  }

  // move cells
  var rangeToMoveStartRow = 4; 
  var rangeToMoveEndRow = 6; 
  var rangeToMoveStartCol = 1; 
  var rangeToMoveEndCol = 8; 

  var rangeToMove = buttonSheet.getRange(rangeToMoveStartRow, rangeToMoveStartCol, rangeToMoveEndRow - rangeToMoveStartRow + 1, rangeToMoveEndCol - rangeToMoveStartCol + 1);
  var valuesToMove = rangeToMove.getValues();
  var numRowsToMove = rangeToMove.getNumRows();
  var numColsToMove = rangeToMove.getNumColumns();

  rangeToMove.breakApart();

  var lastRowButtonSheet = buttonSheet.getLastRow();
  var newRange = buttonSheet.getRange(lastRowButtonSheet + 1, rangeToMoveStartCol, numRowsToMove, numColsToMove);
  newRange.setValues(valuesToMove);

  rangeToMove.clearContent();

  var firstRowToShift = rangeToMoveEndRow + 1;
  var lastRowAfterShift = buttonSheet.getLastRow();
  
  // define first row to shift
  var firstShiftableRow = Math.max(firstRowToShift, 4); // Adjust for the first 3 rows that are occupied

  if (firstShiftableRow <= lastRowAfterShift) {
    var numRowsToShift = lastRowAfterShift - firstShiftableRow + 1;
    if (numRowsToShift > 0) {
      var rowsToShiftUp = buttonSheet.getRange(firstShiftableRow, rangeToMoveStartCol, numRowsToShift, numColsToMove);
      var dataToShift = rowsToShiftUp.getValues();
      
      // place data into gap
      var gapRange = buttonSheet.getRange(firstShiftableRow - 3, rangeToMoveStartCol, numRowsToShift, numColsToMove);
      gapRange.setValues(dataToShift);
    
      // clear rows
      var clearRange = buttonSheet.getRange(lastRowAfterShift - 2, rangeToMoveStartCol, 3, numColsToMove);
      clearRange.clearContent();
    }
  }
  
  //set grey
  var lastDataRow = buttonSheet.getLastRow();
  var rangeToColor = buttonSheet.getRange(lastDataRow - numRowsToMove + 1, rangeToMoveStartCol, numRowsToMove, numColsToMove);
  rangeToColor.setBackground('lightgrey');

  clearRange.clearContent();
  
}
