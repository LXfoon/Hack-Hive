function onFormSubmit(event) {

  record_array = []

  var form = FormApp.openById('14Xky0l0eEzTYnODq6iXTeW-EzWopKL-onLq4nxMb1Pc');
  var formResponses = form.getResponses();
  var formCount = formResponses.length;

  var formResponse = formResponses[formCount - 1];
  var itemResponses = formResponse.getItemResponses();

  for (var j = 0; j < itemResponses.length; j++) {
  var itemResponse = itemResponses[j];
    var title = itemResponse.getItem().getTitle();
    var answer = itemResponse.getResponse();

    Logger.log(title);
    Logger.log(answer);

    record_array.push(answer);
  }
   
  AddRecord(record_array[0], record_array[1], record_array[2], record_array[3], record_array[4]);

}


function addNewRow() {
  var url = 'https://docs.google.com/spreadsheets/d/1bwqkvg1w74ej2qCSiufLmhV5sihRquJDnii3VjYj5G0/edit?gid=1135719921#gid=1135719921';
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Leave Request Approval");
  dataSheet.getRange("A4:H6").copyTo(dataSheet.getRange(dataSheet.getLastRow()+1,1,1,7));
 }

addNewRow();

function AddRecord(name, dolstart, dolend, reason, por) {
  var url = 'https://docs.google.com/spreadsheets/d/1bwqkvg1w74ej2qCSiufLmhV5sihRquJDnii3VjYj5G0/edit?gid=1135719921#gid=1135719921';
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Leave Request Approval");
  var lastRow = dataSheet.getLastRow();
  var blankRow = lastRow + 1;
  var nameRange = dataSheet.getRange(blankRow - 3, 2);
  var dolstartRange = dataSheet.getRange(blankRow + 1 - 3, 2);
  var dolendRange = dataSheet.getRange(blankRow + 1 - 3, 4);
  var reasonRange = dataSheet.getRange(blankRow + 2 - 3, 2);
  var submitdateRange = dataSheet.getRange(blankRow + 2 - 3, 6)
  var porRange = dataSheet.getRange(blankRow - 3, 6)
  nameRange.setValue(name);
  dolstartRange.setValue(dolstart);
  dolendRange.setValue(dolend);
  reasonRange.setValue(reason);
  submitdateRange.setValue(new Date());
  porRange.setValue("View attachment");
}

