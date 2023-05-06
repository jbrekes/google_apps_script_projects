# Report Automation &nbsp; 📊🚀

This project consists of an automation solution for a client's travel booking spreadsheet. The goal was to create a preview of the data in the same spreadsheet, allowing the client to evaluate each trip separately, and to be able to create a new report on demand by specifying the code of a single trip.



## Solution &nbsp; 🛠️

The solution consists of the following steps:

1. A tab was created in the Master Spreadsheet that queries, cleans and formats the data. This provides a preview of how the new document would look like in case the user wants to generate it.

2. An automation using Google Apps Script was developed to create a new Spreadsheet, which links the necessary information from the Master Spreadsheet using functions such as QUERY() and IMPORTRANGE().

3. The result is a new file completely linked to the Master Spreadsheet. This means that any change in a record of the indicated trip is automatically reflected in the new document.

4. Additional controls were implemented to ensure the quality of the output, such as a review of existing documents for the same trip and the verification of the correct format of the data.

## Getting Started &nbsp; 🚀

To use this project, follow these steps:

1. Open a Spreadsheet in your Google Drive account.
2. Go to Extensions -> AppsScript.
3. Copy the code and fill the variables with your IDs.
4. You will have to modify the ranges and formats of the functions depending on the data you want to extract

## Contribution &nbsp; 🤝

If you have any suggestions, feedback, or ideas for improving this project, please feel free to [open an issue](https://github.com/jbrekes/google_apps_script_projects/issues) or [submit a pull request](https://github.com/jbrekes/google_apps_script_projects/pulls).
