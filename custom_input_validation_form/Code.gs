function collectDataFromInputSheet() {
  // Get the 'Input' sheet by name
  var sheetName = 'Input';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Get the range of data starting from cell A3 to column S
  var startRow = 4;
  var startColumn = 1;
  var numRows = sheet.getLastRow() - startRow + 1;
  var numColumns = 19; // Number of columns from A to S
  var range = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var data = range.getValues();

  return data
}

function testing(){
    var valid_ranges = {
      10: [20, 80],  // Lens Height
      11: [20, 80],  // Lens Width
      12: [0, 40],   // Bridge Size
      13: [80, 500], // Frame Width
      14: [0, 500]   // Temple Length
    }

  Logger.log(Object.keys(valid_ranges))
}

function extractPlainText(htmlString) {
  var strippedText = htmlString.replace(/<[^>]+>/g, '');
  return strippedText;
}

function updateErrors(errors_dict, cell, message_to_add){
  if (errors_dict[cell]) {
    // If the cell reference already exists in the dictionary, update the error message
    errors_dict[cell] += '. ' + message_to_add;
  } else {
    // If the cell reference doesn't exist in the dictionary, add it as a new entry
    errors_dict[cell] = message_to_add;
  };
}

function rowValidations(){
  var sheetName = 'Input';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var table = collectDataFromInputSheet();
  var startRow = 4;
  var startColumn = 1;

  var row_errors = {};

  for (var i = 0; i < table.length; i++){
    var row = table[i];
    
    // Mandatory fields
    var mandatory_columns = [0,1,2,3,6,10,11,12,13,14]

    for (var j = 0; j < mandatory_columns.length; j++) {
      var columnIndex = mandatory_columns[j];
      var cellValue = row[columnIndex];

      // Fields must not be empty
      if (cellValue === '') {
        var cell = sheet.getRange(startRow + i, startColumn + columnIndex);
        var cellReference = cell.getA1Notation();

        updateErrors(row_errors, cellReference, 'This field is mandatory, it should not be empty')
      }
      // Fields should have between 1 and 128 characters
      else if (cellValue.length < 1 || cellValue.length > 128) {
        var cell = sheet.getRange(startRow + i, startColumn + columnIndex);
        var cellReference = cell.getA1Notation();

        updateErrors(row_errors, cellReference, 'Text length should be between 1 and 128 characters')        
      }
    }

    // Valid Number Ranges
    var valid_ranges = {
      10: [20, 80],  // Lens Height
      11: [20, 80],  // Lens Width
      12: [0, 40],   // Bridge Size
      13: [80, 500], // Frame Width
      14: [0, 500]   // Temple Length
    }

    var col_list = Object.keys(valid_ranges)

    for (var j = 0; j < col_list.length; j++){
      
      var columnIndex = col_list[j];
      var cellValue = parseInt(row[columnIndex]);
      var rangeMin = valid_ranges[columnIndex][0];
      var rangeMax = valid_ranges[columnIndex][1];

      if (!Number.isInteger(cellValue)){
        var cell = sheet.getRange(startRow + i, startColumn + parseInt(columnIndex));
        var cellReference = cell.getA1Notation();

        updateErrors(row_errors, cellReference, 'Please enter an integer value between ' + String(valid_ranges[columnIndex][0]) + ' and ' + String(valid_ranges[columnIndex][1]))        
      }

      if (cellValue < rangeMin || cellValue > rangeMax) {
        var cell = sheet.getRange(startRow + i, startColumn + parseInt(columnIndex));
        var cellReference = cell.getA1Notation();

        updateErrors(row_errors, cellReference, 'Measures are out of range, please enter a value between ' + String(valid_ranges[columnIndex][0]) + ' and ' + String(valid_ranges[columnIndex][1]))
      }
    }

    // Valid Text Input
    valid_text_dict = {
      7: ['child', 'adolescent', 'adult', 'senior', ''],                    // Age
      8: ['m', 'f', ''],                                                    // Gender
      9: ['round', 'square', 'phantos', 'pilot', 'oval', 'irregular', '']   // Frame Shape
    };

    var text_list = Object.keys(valid_text_dict);

    for (var k = 0; k < text_list.length; k++){
      var columnIndex = text_list[k];
      var cellValue = row[columnIndex].toLowerCase();
      var list_options = valid_text_dict[columnIndex]

      if (!list_options.includes(cellValue)){
        var cell = sheet.getRange(startRow + i, startColumn + parseInt(columnIndex));
        var cellReference = cell.getA1Notation();
        var optionsText = list_options.join(', ');

        updateErrors(row_errors, cellReference, 'You have entered an incorrect value. Please enter one of the following options: ' + optionsText);        
      };
    };

    // Valid HEX color format
    var hexColorPattern = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/; // Regular expression pattern for valid HEX color format
    var columnIndex = 6

    var hexCellValue = row[columnIndex];
    if (!hexColorPattern.test(hexCellValue)){
      var cell = sheet.getRange(startRow + i, startColumn + parseInt(columnIndex));
      var cellReference = cell.getA1Notation();      

      updateErrors(row_errors, cellReference, 'The color code must be entered in HEX format.'); 
    }

    // Valid URLs
    var url_cols = [4,5];
    var urlPattern = /^(http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/)?[a-z0-9]+([\-\.]{1}[a-z0-9]+)*\.[a-z]{2,5}(:[0-9]{1,5})?(\/.*)?$/i;

    for (var j = 0; j < url_cols.length; j++){
      
      var columnIndex = url_cols[j];
      var cellValue = row[columnIndex];

      if (!urlPattern.test(cellValue)){
        var cell = sheet.getRange(startRow + i, startColumn + parseInt(columnIndex));
        var cellReference = cell.getA1Notation();  

        updateErrors(row_errors, cellReference, 'Please write a valid URL.');
      }
    }
  }

  return row_errors

}

function columnValidations(){
  var col_errors = {}

  // Unique Variation Code
  var sheetName = 'Input';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var table = collectDataFromInputSheet();
  var startRow = 4;
  var startColumn = 1;

  var numRows = sheet.getLastRow() - startRow + 1;

  var columnCRange = sheet.getRange(startRow, 4, numRows, 1);
  var columnCValues = columnCRange.getValues();
  
  var uniqueValues = new Set();

  for (var i = 0; i < numRows; i++) {
    var value = columnCValues[i][0];

    if (value !== '') {
      if (uniqueValues.has(value)) {
        var cell = columnCRange.getCell(i + 1, 1);
        var cellReference = cell.getA1Notation(); 

        updateErrors(col_errors, cellReference, 'Duplicated value. The Variation Code must be a unique value');
      } else {
        uniqueValues.add(value);
      }
    }
  }

  // Product ID and Frame Shape combination
  var productIDsRange = sheet.getRange(startRow, 2, numRows, 1);
  var productIDs = productIDsRange.getValues();
  var categories = sheet.getRange(startRow, 10, numRows, 1).getValues();

  var productCategories = {};

  for (var i = 0; i < numRows; i++) {
    var productID = productIDs[i][0];
    var category = categories[i][0];

    if (productID !== '') {
      if (productID in productCategories) {
        if (productCategories[productID] !== category) {
          var cell = productIDsRange.getCell(i + 1, 1);
          var cellReference = cell.getA1Notation(); 

          updateErrors(col_errors, cellReference, 'Each product ID can only be associated with 1 Frame Shape variant.');
        }
      } else {
        productCategories[productID] = category;
      }
    }
  }

  return col_errors

}

function errorMessageBox(){
  var sheetName = 'Input';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var row_errors = rowValidations();
  var col_errors = columnValidations();

  var combinedDict = {};

  // Copy the first dictionary to the combined dictionary
  for (var key in row_errors) {
    combinedDict[key] = row_errors[key];
  }

  // Merge the second dictionary into the combined dictionary
  for (var key in col_errors) {
    if (key in combinedDict) {
      combinedDict[key] += ", " + col_errors[key];
    } else {
      combinedDict[key] = col_errors[key];
    }
  }

  var eror_keys = Object.keys(combinedDict);
  var error_counts = eror_keys.length;

  if (error_counts == 0){
    var full_range = sheet.getRange("A4:S1000");

    full_range.setBackground(null);
    full_range.setFontFamily("Arial");
    full_range.setFontSize(10);

    var errorMessage = '<div style="font-family: Arial;">';
    errorMessage += '<span style="font-size: 12px;">No errors were found in the analyzed information.</span><br><br>';
    errorMessage += '<input type="button" value="OK" onclick="google.script.host.close();" style="float: right; margin-top: 10px; background-color: #3B82F6; color: white; font-size: 14px; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.2);">';
    errorMessage += '</div>';
  } else {
    // Order the combined dictionary based on keys in ascending order
    var sortedDict = {};
    var sortedKeys = Object.keys(combinedDict).sort();

    for (var i = 0; i < sortedKeys.length; i++) {
      var key = sortedKeys[i];
      sortedDict[key] = combinedDict[key];
    }

    var errorMessage = '<div style="font-family: Arial;">';
    errorMessage += '<span style="font-size: 12px;">We found the following errors, please verify that the information entered complies with the following validations. Cells with errors have been highlighted correctly on the sheet to facilitate their location.</span><br><br>';

    for (var cell in sortedDict) {
      var range = sheet.getRange(cell);
      range.setBackground("#fd33a4");

      errorMessage += "<b style='font-size: 14px;'>" + cell + "</b><br>";
      errorMessage += "<span style='font-size: 12px;'>" + sortedDict[cell] + "</span><br><br>";
    }

    errorMessage = errorMessage.trim();

    errorMessage += '<input type="button" value="OK" onclick="google.script.host.close();" style="float: right; margin-top: 10px; background-color: #3B82F6; color: white; font-size: 14px; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.2);">';
    errorMessage += '</div>';
  }
  
  var htmlOutput = HtmlService.createHtmlOutput(errorMessage);
  htmlOutput.setTitle("Error Report");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Error Report");

  // Paste the same error message in a cell for reference
  // var targetCell = 'F1';
  // var targetRange = sheet.getRange(targetCell);
  // var plainTextMessage = extractPlainText(errorMessage);

  // targetRange.setValue(plainTextMessage);

  resetValidations()

  if (error_counts == 0){
    return true
  } else {
    return false
  }
}

function downloadCsv() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("A3:S1000");
  var validations = errorMessageBox();

  if (validations) {
    var dataValues = dataRange.getValues();
    var csvContent = dataValues.map(row => row.join(',')).join('\n');
    var fileName = "exported_data.csv";
    var fileBlob = Utilities.newBlob(csvContent, MimeType.CSV, fileName);

    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Save CSV File', 'Please specify the file name:', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
      var userFileName = response.getResponseText();
      fileBlob.setName(userFileName + ".csv");

      var fileUrl = DriveApp.createFile(fileBlob).getDownloadUrl();
      var htmlOutput = '<div style="text-align: center; font-family: Arial;">' +
        '<h2>Download CSV File</h2>' +
        '<p>Click the button below to download the CSV file:</p>' +
        '<a href="' + fileUrl + '" download="' + userFileName + '.csv">' +
        '<button style="padding: 10px 20px; font-size: 16px; background-color: #3B82F6; color: white; font-size: 14px; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.2)">Download</button></a></div>';

      var userInterface = HtmlService.createHtmlOutput(htmlOutput)
        .setWidth(400)
        .setHeight(150);
      
      ui.showModalDialog(userInterface, "Download CSV File");
    }
  }
}

function resetValidations(){
  var sheetName = 'Input';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Text Length
  var tl_range = sheet.getRange("A4:C1000");
  
  var tl_rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=AND(LEN(A4) < 128,LEN(A4) > 1)')
    .setHelpText("The text has exceeded the maximum value of 128 characters.")
    .setAllowInvalid(false)
    .build();
  
  tl_range.setDataValidation(tl_rule);  

  // Unique Variation Code
  var vc_range = sheet.getRange("D4:D1000");
  
  var vc_rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=AND(COUNTIF($D3:$D,"="&C4)  = 1, LEN(C4) < 128,LEN(C4) > 1)')
    .setHelpText("Duplicated value. The Variation Code must be a unique value.")
    .setAllowInvalid(false)
    .build();
  
  vc_range.setDataValidation(vc_rule);  

  // Valid URLs
  var url_range = sheet.getRange("E4:F1000");

  var url_rule = SpreadsheetApp.newDataValidation()
    .requireTextIsUrl()
    .setAllowInvalid(false)
    .setHelpText("The text must be a valid URL.")
    .build();

  url_range.setDataValidation(url_rule);

  // Valid HEX Color Code
  var color_range = sheet.getRange("G4:G1000");

  var color_rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=REGEXMATCH(G4,"^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$")')
    .setHelpText("The color code must be entered in HEX format.")
    .setAllowInvalid(false)
    .build();  

  color_range.setDataValidation(color_rule);

  // Age options
  var age_range = sheet.getRange("H4:H1000");

  var age_rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Adolescent', 'Adult','Child','Senior'])
    .setHelpText("The age, if filled in, must be one of the following values: Child, Adolescent, Adult, Senior.")
    .setAllowInvalid(false)
    .build();  

  age_range.setDataValidation(age_rule);

  // Sex options
  var sex_range = sheet.getRange("I4:I1000");

  var sex_rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['F', 'M'])
    .setHelpText("The age, if filled in, must be one of the following values: F, M.")
    .setAllowInvalid(false)
    .build();  

  sex_range.setDataValidation(sex_rule);

  // Frame Shape options
  var frame_range = sheet.getRange("J4:J1000");

  var frame_rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Irregular', 'Oval','Phantos','Pilot', 'Round', 'Square'])
    .setHelpText("Frame Shape, if filled in, must be one of the following values: Round, Square, Phantos, Pilot, Oval, Irregular.")
    .setAllowInvalid(false)
    .build();  

  frame_range.setDataValidation(frame_rule);   

  // Valid Numer Ranges
  var num_ranges = [
    ["K4:L1000", 20, 80, "Lens height and width must be between 20 and 80 (measured in mm). Please write only numbers."],
    ["M4:M1000", 0, 40, "Bridge size must be between 0 and 40 (measured in mm). Please write only numbers."],
    ["N4:N1000", 80, 500, "Frame width must be between 80 and 500 (measured in mm). Please write only numbers."],
    ["O4:O1000", 0, 500, "Temple length must be between 0 and 500 (measured in mm). Please write only numbers."]
  ];

  for (var i = 0; i < num_ranges.length; i++){
    var record = num_ranges[i];

    var num_range = sheet.getRange(record[0]);

    var num_rule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(record[1], record[2])
      .setHelpText(record[3])
      .setAllowInvalid(false)
      .build();  

    num_range.setDataValidation(num_rule);    
  }
}

function resetSpreadsheet(){
  var sheetName = 'Input';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  resetValidations();

  var full_range = sheet.getRange("A4:S1000");

  full_range.setBackground(null);
  full_range.setFontFamily("Arial");
  full_range.setFontSize(10);
}