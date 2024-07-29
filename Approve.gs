function approveAndMoveRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var buttonSheet = ss.getSheetByName('Leave Request Approval'); 
  var targetSheet = ss.getSheetByName('2024'); 
  
  if (!buttonSheet || !targetSheet) {
    Logger.log('One or both sheets not found.');
    return;
  }

  // read date range frm buttonSheet
  var startDate = new Date(buttonSheet.getRange('B5').getValue());
  var endDate = new Date(buttonSheet.getRange('D5').getValue());

  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    Logger.log('Invalid date range specified.');
    return;
  }

  // update status in the employee attandence's sheet
  var dateColumnIndex = 2;
  var statusColumnIndex = 11; 
  var lastRow = targetSheet.getLastRow();
  var dateRange = targetSheet.getRange(1, dateColumnIndex, lastRow);
  var values = dateRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var cellDate = new Date(values[i][0]);

    if (cellDate >= startDate && cellDate <= endDate) {
      var row = i + 1; // 1-based row index
      var statusCell = targetSheet.getRange(row, statusColumnIndex);

      // check if merged
      var mergedRanges = statusCell.getMergedRanges();
      if (mergedRanges.length > 0) {
        mergedRanges.forEach(function(range) {
          range.setValue("ON LEAVE");
        });
      } else {
        statusCell.setValue("ON LEAVE");
      }

      // merge with next column
      targetSheet.getRange(row, statusColumnIndex, 1, 2).mergeAcross();
    }
  }

  // move data
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
  var firstShiftableRow = Math.max(firstRowToShift, 3); 

  if (firstShiftableRow <= lastRowAfterShift) {
    var numRowsToShift = lastRowAfterShift - firstShiftableRow + 1;
    if (numRowsToShift > 0) {
      var rowsToShiftUp = buttonSheet.getRange(firstShiftableRow, rangeToMoveStartCol, numRowsToShift, numColsToMove);
      var dataToShift = rowsToShiftUp.getValues();
      
      // place shifted data into gap
      var gapRange = buttonSheet.getRange(firstShiftableRow - 3, rangeToMoveStartCol, numRowsToShift, numColsToMove);
      gapRange.setValues(dataToShift);
      
      // clear duplicated rows
      var clearRange = buttonSheet.getRange(lastRowAfterShift - 2, rangeToMoveStartCol, 3, numColsToMove);
      clearRange.clearContent();
    }
  }

  // set grey
  var lastDataRow = buttonSheet.getLastRow();
  var rangeToColor = buttonSheet.getRange(lastDataRow - numRowsToMove + 1, rangeToMoveStartCol, numRowsToMove, numColsToMove);
  rangeToColor.setBackground('lightgrey');

  clearRange.clearContent();
  
}
