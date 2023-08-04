<h1 align="center">Google Apps Script - Data Validation for Google Sheets</h1>

<p align="center">
  <img src="https://img.shields.io/badge/language-Google%20Apps%20Script-yellow.svg" alt="Language">
  <img src="https://img.shields.io/badge/platform-Google%20Sheets-blue.svg" alt="Platform">
</p>

<p align="center">A collection of powerful data validation functions for your Google Sheets.</p>

---

## Overview

This Google Apps Script code provides robust data validation functionalities for a Google Sheet named 'Input'. The primary purpose of these validations is to safeguard against the limitations or problems of working with validations directly in Google Sheets. One of the main challenges arises when a client copies and pastes information from another place, which often leads to overwriting the validations. As a result, the spreadsheet can quickly lose its usefulness and integrity.

The code addresses these issues by reapplying data validations and ensuring that the 'Input' spreadsheet can always return to its initial state, even after data is copied and pasted. It offers a collection of powerful functions that perform various checks, including mandatory field checks, valid number ranges, valid text inputs, HEX color format validation, and valid URLs.

By using this code, you can maintain the data integrity of your 'Input' sheet, minimize manual data entry errors, and enhance the overall usability of the spreadsheet for all users.

## Functions and Their Functionalities

#### `collectDataFromInputSheet()`

Retrieves data from the 'Input' sheet starting from cell A4 to column S and returns the data as a two-dimensional array.

#### `extractPlainText(htmlString)`

Takes an HTML string and removes all HTML tags, returning only the plain text content.

#### `updateErrors(errors_dict, cell, message_to_add)`

Updates the `errors_dict` object with error messages for a specific `cell`.

#### `rowValidations()`

Performs various validations on each row of data in the 'Input' sheet. It checks for mandatory fields, valid number ranges, valid text inputs, valid HEX color format, and valid URLs. Any errors found are stored in the `row_errors` object.

#### `columnValidations()`

Performs additional validations on specific columns in the 'Input' sheet. It checks for unique Variation Code values and ensures that each Product ID is associated with only one Frame Shape variant. Any errors found are stored in the `col_errors` object.

### `errorMessageBox()`

Combines the errors found in `row_errors` and `col_errors` objects, displays the errors in a user-friendly HTML modal dialog, and highlights the cells with errors on the 'Input' sheet.

#### `downloadCsv()`

Triggers the error message box and, if there are no errors, allows the user to download the data in the 'Input' sheet as a CSV file.

#### `resetValidations()`

Resets data validations for specific ranges in the 'Input' sheet to their default settings. It also removes cell background colors and restores the font style and size.

#### `resetSpreadsheet()`

Resets the entire 'Input' sheet by calling `resetValidations()` and removing all cell background colors while preserving the font style and size.

## How to Use

1. Copy the provided Google Apps Script code to your Google Sheets.
2. Save the script, and it will be bound to the 'Input' sheet.
3. Run the `errorMessageBox()` function to perform data validations and display any errors found.
4. If there are no errors, the `downloadCsv()` function allows you to download the data in the 'Input' sheet as a CSV file.

Please note that this code is designed specifically for the "Input" sheet and for specific ranges of cells. Make sure that your sheet has the appropriate structure and column names to work with this code effectively, or adapt these ranges according to your needs.
Happy coding!

---

<p align="center">Made with ❤️ by Juan Brekes</p>
