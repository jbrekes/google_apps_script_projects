const FOLDERID = 'YOUR_FOLDER_ID'; // The folder ID where the generated reports will be stored
const MASTERFILEID = 'YOUR_MASTER_SPREADSHEET_ID'; //  The ID of the master file spreadsheet containing the source data to which the projects will be linked

function checkExistingFiles() {
  var folder = DriveApp.getFolderById(FOLDERID);
  var files = folder.getFiles();
  var fileNames = [];
  
  while (files.hasNext()) {
    var file = files.next();
    fileNames.push(file.getName());
  }

  return fileNames;
}

function checkExistingURLs() {
  var folder = DriveApp.getFolderById(FOLDERID);
  var files = folder.getFiles();
  var fileURLs = [];
  
  while (files.hasNext()) {
    var file = files.next();
    fileURLs.push(file.getUrl());
  }
  return fileURLs;
}

function getFlightCode() {
  var sheet = SpreadsheetApp.openById(MASTERFILEID).getSheetByName("Report Preview"); 
  var cell = sheet.getRange("B2"); 
  var value = cell.getValue().toUpperCase();
  
  return value;
}

function copyTitles(newSheetId, tabName, titleRange) {
  // Get the source and destination spreadsheets and sheets
  var sourceSpreadsheet = SpreadsheetApp.openById(MASTERFILEID);
  var sourceSheet = sourceSpreadsheet.getSheetByName("Report Preview");
  var destinationSpreadsheet = SpreadsheetApp.openById(newSheetId);
  var destinationSheet = destinationSpreadsheet.getSheetByName(tabName);

  // Define the source range to copy
  var sourceRange = sourceSheet.getRange(titleRange);

  // Get the size of the source range
  var numRows = sourceRange.getNumRows();
  var numColumns = sourceRange.getNumColumns();

  // Get the values and number formats from the source range
  var values = sourceRange.getValues();
  var numberFormats = sourceRange.getNumberFormats();

  // Define the destination range to copy to (A1 in this example)
  var destinationRange = destinationSheet.getRange(1, 1, numRows, numColumns);

  // Set the values and number formats of the destination range to match the source range
  destinationRange.setValues(values);
  destinationRange.setNumberFormats(numberFormats);
}

function formatTitles(sheetId, tabName, titleRange, subTitleRange){
  var sheet = SpreadsheetApp.openById(sheetId);
  var destinationSheet = sheet.getSheetByName(tabName);

  // Modify title Color
  var titleRange = destinationSheet.getRange(titleRange);
  titleRange.setBackground("#bbd2ab");
  titleRange.setFontWeight("bold");
  var subtitleRange = destinationSheet.getRange(subTitleRange);
  subtitleRange.setBackground("#efefef");
}

function addGuestInformation(sheetName, sheetId, tabName){
  var sheet = SpreadsheetApp.openById(sheetId);
  var GuestSheet = sheet.getSheetByName(tabName);
  var cell = GuestSheet.getRange("A3");
  var likeAux = "'%" + sheetName + "%'";
  var passPortAux = "";
  var formula = cell.setFormula(
    '=QUERY({IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!A:C"),ARRAYFORMULA("First Name: " & IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AF:AF") & CHAR(10) & "Last Name " & IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AG:AG") & CHAR(10) & "Nationality: " & IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AD:AD") & CHAR(10) & "Passport Nr: " & IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AC:AC") & CHAR(10) & IF(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AE:AE") = "","","Passport Exp: " & TEXT(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AE:AE"),"YYYY-MM-DD"))),ARRAYFORMULA("Status: " & IF(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!Q:Q") = "na","",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!Q:Q")) & CHAR(10) & "Notes: " & IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!R:R")),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AH:AI"),ARRAYFORMULA(SWITCH(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T"),FALSE,"N",TRUE,"Y",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T"))),ARRAYFORMULA(SWITCH(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!G:G"),FALSE,"N",TRUE,"Y",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!G:G"))),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!J:J"),ARRAYFORMULA(IF(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!I:I")="","",TEXT(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!I:I"),"DD MMM") & " " & TEXT(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!K:K"),"HH:MM"))),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!O:O"),ARRAYFORMULA(IF(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!M:M")="","",TEXT(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!M:M"),"DD MMM") & " " & TEXT(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!N:N"),"HH:MM"))),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AN:AN")},"SELECT Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9,Col14,Col10,Col11,Col12,Col13 WHERE Col1 LIKE ' + likeAux + '",0)'
  )

  // Arrange some columns
  GuestSheet.getRange("A:Z").setVerticalAlignment("middle").setWrap(true);
  GuestSheet.setColumnWidths(3,4,200);
  GuestSheet.setColumnWidth(2,300);
  GuestSheet.setColumnWidth(11,100);
  GuestSheet.setColumnWidth(13,100);

  // Set Conditional Format Rule
  var conditionalRange1 = GuestSheet.getRange("A3:L");
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$G3="Y"')
                           .setBackground("#e8d4dc")
                           .setRanges([conditionalRange1])
                           .build();

  var rules = GuestSheet.getConditionalFormatRules();

  rules.push(rule1);
  GuestSheet.setConditionalFormatRules(rules);

  var conditionalRange2 = GuestSheet.getRange("H3:H");
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$H3="Y"')
                           .setBackground("#c8dccc")
                           .setRanges([conditionalRange2])
                           .build();

  rules.push(rule2);
  GuestSheet.setConditionalFormatRules(rules);

  var conditionalRange3 = GuestSheet.getRange("I3:I");
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$I3="Y"')
                           .setBackground("#c8dccc")
                           .setRanges([conditionalRange3])
                           .build();

  rules.push(rule3);
  GuestSheet.setConditionalFormatRules(rules);
}

function addDietaries(sheetName, sheetId, tabName){
  var sheet = SpreadsheetApp.openById(sheetId);
  var DietariesSheet = sheet.getSheetByName(tabName);
  var cell = DietariesSheet.getRange("A3");
  var likeAux = "'%" + sheetName + "%'"; 

  var formula = cell.setFormula(
    '=QUERY({IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!A:C"),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AH:AH"),ARRAYFORMULA(SWITCH(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T"),FALSE,"N",TRUE,"Y",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T")))},"SELECT Col2,Col3,Col4,Col5 WHERE Col1 LIKE ' + likeAux + '",0)'
  )

  // Arrange some columns
  DietariesSheet.getRange("A:Z").setVerticalAlignment("middle").setWrap(true);
  DietariesSheet.setColumnWidth(1,200);
  DietariesSheet.setColumnWidth(2,300);
  DietariesSheet.setColumnWidth(3,400);

  // Set Conditional Format Rule
  var conditionalRange = DietariesSheet.getRange("A3:D");
  var rule = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$D3="Y"')
                           .setBackground("#e8d4dc")
                           .setRanges([conditionalRange])
                           .build();

  var rules = DietariesSheet.getConditionalFormatRules();

  // Hide Canceled Column
  var hideTitle = DietariesSheet.getRange("D2");
  hideTitle.setValue("Canceled");
  var hideRange = DietariesSheet.getRange("D:D");
  DietariesSheet.hideColumn(hideRange);

  rules.push(rule);
  DietariesSheet.setConditionalFormatRules(rules);
}

function addEmercencyDetails(sheetName, sheetId, tabName){
  var sheet = SpreadsheetApp.openById(sheetId);
  var EmergencySheet = sheet.getSheetByName(tabName);
  var cell = EmergencySheet.getRange("A3");
  var likeAux = "'%" + sheetName + "%'"; 

  var formula = cell.setFormula(
    '=QUERY({IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!A:C"),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!AJ:AM"),ARRAYFORMULA(SWITCH(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T"),FALSE,"N",TRUE,"Y",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T")))},"SELECT Col2,Col3,Col4,Col5,Col6,Col7,Col8 WHERE Col1 LIKE ' + likeAux + '",0)'
  )

  // Arrange some columns
  EmergencySheet.getRange("A:Z").setVerticalAlignment("middle").setWrap(true);
  EmergencySheet.setColumnWidth(1,200);
  EmergencySheet.setColumnWidth(2,300);

  // Set Conditional Format Rule
  var conditionalRange = EmergencySheet.getRange("A3:G");
  var rule = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$G3="Y"')
                           .setBackground("#e8d4dc")
                           .setRanges([conditionalRange])
                           .build();

  var rules = EmergencySheet.getConditionalFormatRules();

  // Hide Canceled Column
  var hideTitle = EmergencySheet.getRange("G2");
  hideTitle.setValue("Canceled");
  var hideRange = EmergencySheet.getRange("G:G");
  EmergencySheet.hideColumn(hideRange);

  rules.push(rule);
  EmergencySheet.setConditionalFormatRules(rules);
}

function addInsuranceDetails(sheetName, sheetId, tabName){
  var sheet = SpreadsheetApp.openById(sheetId);
  var InsuranceSheet = sheet.getSheetByName(tabName);
  var cell = InsuranceSheet.getRange("A3");
  var likeAux = "'%" + sheetName + "%'"; 

  var formula = cell.setFormula(
    '=QUERY({IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!A:C"),IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!V:Y"),ARRAYFORMULA(SWITCH(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T"),FALSE,"N",TRUE,"Y",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T")))},"SELECT Col2,Col3,Col4,Col5,Col6,Col7,Col8 WHERE Col1 LIKE ' + likeAux + '",0)'
  )

  // Arrange some columns
  InsuranceSheet.getRange("A:Z").setVerticalAlignment("middle").setWrap(true);
  InsuranceSheet.setColumnWidths(3,4,200);
  InsuranceSheet.setColumnWidth(1,200);
  InsuranceSheet.setColumnWidth(2,300);
  InsuranceSheet.setColumnWidth(6,300);

  // Set Conditional Format Rule
  var conditionalRange = InsuranceSheet.getRange("A3:G");
  var rule = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$G3="Y"')
                           .setBackground("#e8d4dc")
                           .setRanges([conditionalRange])
                           .build();

  var rules = InsuranceSheet.getConditionalFormatRules();

  // Hide Canceled Column
  var hideTitle = InsuranceSheet.getRange("G2");
  hideTitle.setValue("Canceled");
  var hideRange = InsuranceSheet.getRange("G:G");
  InsuranceSheet.hideColumn(hideRange);

  rules.push(rule);
  InsuranceSheet.setConditionalFormatRules(rules);
}

function addRoomAllocations(sheetName, sheetId, tabName){
  var sheet = SpreadsheetApp.openById(sheetId);
  var RoomSheet = sheet.getSheetByName(tabName);
  var cell = RoomSheet.getRange("A3");
  var likeAux = "'%" + sheetName + "%'"; 

  var formula = cell.setFormula(
    '=QUERY({IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!A:E"),ARRAYFORMULA(SWITCH(IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T"),FALSE,"N",TRUE,"Y",IMPORTRANGE("' + MASTERFILEID + '","Master_Sheet_Temp!T:T")))},"SELECT Col2,Col3,Col5,Col6 WHERE Col1 LIKE ' + likeAux + '",0)'
  )

  // Arrange some columns
  RoomSheet.getRange("A:Z").setVerticalAlignment("middle").setWrap(true);
  RoomSheet.setColumnWidths(1,3,200);

  // Set Conditional Format Rule
  var conditionalRange = RoomSheet.getRange("A3:D");
  var rule = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$D3="Y"')
                           .setBackground("#e8d4dc")
                           .setRanges([conditionalRange])
                           .build();

  var rules = RoomSheet.getConditionalFormatRules();

  // Hide Canceled Column
  var hideTitle = RoomSheet.getRange("D2");
  hideTitle.setValue("Canceled");
  var hideRange = RoomSheet.getRange("D:D");
  RoomSheet.hideColumn(hideRange);

  rules.push(rule);
  RoomSheet.setConditionalFormatRules(rules);
}

function addExtras(sheetName, sheetId, tabName){
  var sheet = SpreadsheetApp.openById(sheetId);
  var ExtrasSheet = sheet.getSheetByName(tabName);
  var cell = ExtrasSheet.getRange("A3");
  var likeAux = "'%" + sheetName + "%'"; 


  var formula = cell.setFormula(
    '=QUERY({IMPORTRANGE("' + MASTERFILEID + '","Extras_Clean!A:S"),ARRAYFORMULA(IF(TEXT(IMPORTRANGE("' + MASTERFILEID + '","Extras_Clean!G:H"),"YYYY-MM-DD") = "1899-12-30","",IMPORTRANGE("' + MASTERFILEID + '","Extras_Clean!G:H"))),ARRAYFORMULA(IF(TEXT(IMPORTRANGE("' + MASTERFILEID + '","Extras_Clean!L:M"),"YYYY-MM-DD") = "1899-12-30","",IMPORTRANGE("' + MASTERFILEID + '","Extras_Clean!L:M")))},"SELECT Col3,Col4,Col5,Col9,Col20,Col21,Col10,Col11,Col14,Col22,Col23,Col15,Col16,Col17,Col18,Col19 WHERE Col1 LIKE ' + likeAux + '",0)'
  )

  // Arrange some columns
  ExtrasSheet.getRange("A:Z").setVerticalAlignment("middle").setWrap(true);
  ExtrasSheet.setColumnWidth(2,300);
  ExtrasSheet.setColumnWidth(7,300);
  ExtrasSheet.setColumnWidth(12,300);
  ExtrasSheet.setColumnWidth(14,300);
  ExtrasSheet.setColumnWidth(15,300);

  // Set Conditional Format Rule
  var conditionalRange = ExtrasSheet.getRange("A3:P");
  var rule = SpreadsheetApp.newConditionalFormatRule()
                           .whenFormulaSatisfied('=$P3="Y"')
                           .setBackground("#e8d4dc")
                           .setRanges([conditionalRange])
                           .build();

  var rules = ExtrasSheet.getConditionalFormatRules();

  // Hide Canceled Column
  var hideTitle = ExtrasSheet.getRange("P2");
  hideTitle.setValue("Canceled");
  var hideRange = ExtrasSheet.getRange("P:P");
  ExtrasSheet.hideColumn(hideRange);

  rules.push(rule);
  ExtrasSheet.setConditionalFormatRules(rules);
}

function createNewFile(){

  // Check if the file already exists
  var folder = DriveApp.getFolderById(FOLDERID);
  var existingFiles = checkExistingFiles();
  var existingURLs = checkExistingURLs();
  var name = getFlightCode();

  // If the file doesn't exist, create it and paste the info
  if (existingFiles.includes(name) == false){

    // Create the file and get it's ID
    var newSpreadsheet = SpreadsheetApp.create(name);
    var fileId = newSpreadsheet.getId();
    var file = DriveApp.getFileById(fileId);

    // Add new tabs and name them. Paste Titles from Master Sheet
    var ss = SpreadsheetApp.openById(fileId);

    var tabs = ['Guest Details', 'Dietaries', 'Emergency Details', 'Insurance Details', 'Room Allocations','Extras'];
    var masterRanges = ["A6:M7","O6:Q7",'S6:X7','Z6:AE7','AG6:AI7','AK6:AZ7'];
    var titleRanges = ["A1:M1","A1:C1",'A1:F1','A1:F1','A1:C1','A1:O1'];
    var subTitleRanges = ["A2:M2","A2:C2",'A2:F2','A2:F2','A2:C2','A2:O2'];

    for (var i = 0; i < tabs.length; i ++) {
      var tabName = tabs[i];
      var newSheet = ss.insertSheet();
      newSheet.setName(tabName);

      copyTitles(fileId,tabs[i],masterRanges[i]);
      formatTitles(fileId,tabs[i],titleRanges[i],subTitleRanges[i]);

      newSheet.autoResizeColumns(1,15);
    }

    var sheetToDelete = ss.getSheetByName("Sheet1");
    ss.deleteSheet(sheetToDelete);

    // Add Data
    addGuestInformation(name,fileId,'Guest Details');
    addDietaries(name,fileId,'Dietaries');
    addEmercencyDetails(name,fileId,'Emergency Details');
    addInsuranceDetails(name,fileId,'Insurance Details');
    addRoomAllocations(name,fileId,'Room Allocations');
    addExtras(name,fileId,'Extras');

    // Move file to desired folder and remove it from Root
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    var sheetURL = file.getUrl();

    // Show Success message and button to access the new report
    var message = 'The report was successfully created. File Name: ' + name + '\n \nPress "OK" to go to the file';
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Report Created', message, ui.ButtonSet.OK_CANCEL);
    if (response == ui.Button.OK) {
      var html = '<script> window.open("' + sheetURL + '"); google.script.host.close(); </script>';
      ui.showModalDialog(HtmlService.createHtmlOutput(html), 'Existing File');
    }

    return fileId;
  } else {
    var nameIndex = existingFiles.indexOf(name);
    var nameURL = existingURLs[nameIndex];
    var message = 'It appears that the Trip Code you are trying to use already has a file created. If you want to create it again, delete the previous file. \n \nPress "OK" to go to the file';
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('File Already Exists', message, ui.ButtonSet.OK_CANCEL);
    if (response == ui.Button.OK) {
      var html = '<script> window.open("' + nameURL + '"); google.script.host.close(); </script>';
      ui.showModalDialog(HtmlService.createHtmlOutput(html), 'Existing File');
    }
  }
}
