var masterSpreadsheetId = 'YOUR_MASTER_SPREADSHEET_ID'; 

function GetGoogleNewsCounts(search, monthAndYear) {
    if(search != "" && monthAndYear != ""){
        const url = 'CLOUD_FUNCTIONS_GOOGLE_SEARCH_APP_URL';
        const options = {
        "method": "post",
        "payload": {
            "search": search,
            "monthAndYear": monthAndYear
        }
        };
        const response = tryGetInfo(url, options);
        return Number(JSON.parse(response.getContentText()).resultsNumber);
    }
    else{
        return "";
    }
}
  
function GetBingNewsCounts(search, monthAndYear) {
    if(search != "" && monthAndYear != ""){
        const url = 'CLOUD_FUNCTIONS_BING_SEARCH_APP_URL';
        const options = {
        "method": "post",
        "payload": {
            "search": search,
            "monthAndYear": monthAndYear
        }
        };
        const response = tryGetInfo(url, options);
        return Number(JSON.parse(response.getContentText()).resultsNumber);
    }
    else{
        return "";
    }
}
  
function tryGetInfo(url, options){
    const response = UrlFetchApp.fetch(url, options);
    return response;
}
  
function getMonthIndex(monthName) {
    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return months.indexOf(monthName);
}
  
function updateBDD() {
    var sheet = SpreadsheetApp.openById('masterSpreadsheetId').getSheetByName('Main Search'); 
    var search_range = sheet.getRange('C10');
    var search_value = search_range.getValues()[0];
    var country_range = sheet.getRange('C11');
    var country_value = country_range.getValues()[0];
    var search_engine_range = sheet.getRange('C12');
    var search_engine_value = search_engine_range.getValues()[0];
    var result_range = sheet.getRange('C16:O16');
    var result_values = result_range.getValues();
    var dates_range = sheet.getRange('C15:O15');
    var dates_values = dates_range.getValues();

    // Filter out any null values from the array
    result_values = result_values[0].filter(function(value) {
        return value !== '';});
    dates_values = dates_values[0].filter(function(value) {
        return value !== '';});

    // Create a date from the dates_values field

    var month_searched_start_date_lst = [];

    for (var i = 0; i < dates_values.length; i ++){
        var date_temp = dates_values[i];

        var month_string = date_temp.substring(0, 3);
        var month_value = getMonthIndex(month_string);

        var year_value = parseInt(date_temp.substring(date_temp.length - 2));
        if(year_value < 50){
        year_value = year_value + 2000;
        }
        else{
        year_value = year_value + 1900;
        }

        var date_to_add = new Date(year_value, month_value, 1);
        date_to_add = Utilities.formatDate(date_to_add, "GMT", "yyyy-MM-dd");
        
        month_searched_start_date_lst.push(date_to_add);
    }

    // Create a list with the generated keys for the new search

    var search_keys_lst = [];

    for (var i = 0; i < result_values.length; i++) {
        var search_key = country_value[0] + search_value[0] + dates_values[i] + search_engine_value[0].substring(0, 3)

        search_key = search_key.toUpperCase().replace(/\s/g, "");

        search_keys_lst.push(search_key);
    }

    // Get all saved ids

    var bdd_sheet = SpreadsheetApp.openById(masterSpreadsheetId).getSheetByName('BDD');

    var saved_ids = bdd_sheet.getRange('B:B')
    var saved_ids_values = saved_ids.getValues()

    var saved_ids_values_filtered = []

    for (var i = 0; i < saved_ids_values.length; i++){
        if(saved_ids_values[i][0] !== ''){
        saved_ids_values_filtered.push(saved_ids_values[i][0])
        }
    }
    saved_ids_values = saved_ids_values.filter(function(value) {
        return value !== '' && value !== null && value.length > 0;});

    // Look if a searched value is already in the BDD
    var commonElements = [];

    for (var i = 0; i < search_keys_lst.length; i++) {
        if (saved_ids_values_filtered.includes(search_keys_lst[i])) {
        commonElements.push(search_keys_lst[i]);
        }
    }

    // Look for the indices of the rows that I need to delete

    var bdd_to_delete = bdd_sheet.getDataRange();
    var bdd_values = bdd_to_delete.getValues();

    var bdd_indexes = [];

    for (var i = 0; i < bdd_values.length; i++) {
        var id = bdd_values[i][1]; // The ID is in the second column
        if (commonElements.indexOf(id) !== -1) { // check if the ID is present in the list
        bdd_indexes.push(i + 1); // i+1 because row indices start at 1
        }
    }

    Logger.log(bdd_indexes)

    if(bdd_indexes.length > 0){
        for (var i = bdd_indexes.length - 1; i >= 0; i --){
        Logger.log(bdd_indexes[i])
        bdd_sheet.deleteRow(bdd_indexes[i])
        }
    }

    // Add N rows at the start of the table containing the new search terms

    for (var i = 0; i < result_values.length; i++){
        bdd_sheet.insertRowAfter(1);

        var currentDate = new Date();
        var formattedDate = Utilities.formatDate(currentDate, "GMT", "yyyy-MM-dd");

        var row_to_add = [formattedDate, search_keys_lst[i], search_engine_value[0], search_value[0], country_value[0], dates_values[i], month_searched_start_date_lst[i], result_values[i]];

        Logger.log(row_to_add)

        bdd_sheet.getRange('A2:H2').setValues([row_to_add]);

    }

    return result_values; 
}
