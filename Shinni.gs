// Copy Google Form Requests to Leave Request Approval Page
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gform = ss.getSheetByName('Leave Form'); 
  var leaveapp = ss.getSheetByName('2024')

function test() {
  var lastRow = gform.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  Logger.log(lastRow).string;
}

// attempt#2

function copyshit() {
  var url = 'https://docs.google.com/spreadsheets/d/1bwqkvg1w74ej2qCSiufLmhV5sihRquJDnii3VjYj5G0/edit?gid=1135719921#gid=1135719921';
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Leave Request Approval");
  dataSheet.getRange("A4:H6").copyTo(dataSheet.getRange(dataSheet.getLastRow()+1,1,1,7));
 }

 copyshit();