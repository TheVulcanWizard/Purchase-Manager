
function archiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var sheet_name = sheet.getName()
  var last_row = sheet.getLastRow()  
  
  for(var i = 2; i <= last_row; i++) {
    var completed_box = sheet.getRange(i, 9).getValue()
    if(completed_box == false) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert("You are attempting to archive a sheet without sending every item. Are you sure you want to continue?", ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) {
        break;
      }
      else {
        return;
      }
    }
  }
  
  var destination = SpreadsheetApp.openById("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
  var new_sheet = sheet.copyTo(destination)
  new_sheet.setName(sheet_name)
  SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet()
}