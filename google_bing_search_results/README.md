# 🌐 Web Scraping and Google Sheets Automation

This project consists of a web scraping automation tool that searches for specific search terms and timeframes in both Google News and Bing search engines. The search results are then stored in a Google Sheet for later use. 

This repository contains the automation done in AppsScript, **not** the web scrapping process or the creation of the Cloud Functions.

## 🔧 Tools Used

- Node.js Puppeteer Library
- Google Cloud Function
- Google Sheets
- AppsScript

## 🚀 How it Works

The web scraping automation was created using the Node.js Puppeteer library. We created a Google Cloud Function to wrap the automation and allow it to be integrated with Google Sheets. 

The final function searches for a specific search term and timeframe in both Google News and Bing search engines, and retrieves the estimated total search results. The results are then automatically stored in a Google Sheet using AppsScript to automate the search and storing process. 

## 📷 Screenshots

Here are some screenshots of the final product:

**Main Search:** You must select the term to search for, the search engine and the months for which you want to obtain the result. Previous saved searches related to that search term will be displayed at the bottom.
In case you click on the "Save Search" button, the previous search records matching the current search conditions (search term, search engine and time frame) will be deleted and replaced by the current ones. In addition, the new results will be added to the record

![Main Search](https://github.com/jbrekes/google_apps_script_projects/blob/main/google_bing_search_results/Main%20Search.png)

**Saved Searches:** The previous searches performed are stored here. Additional information such as the date the search was performed is attached, as well as a unique ID for easy identification.

![Saved Searches](https://github.com/jbrekes/google_apps_script_projects/blob/main/google_bing_search_results/Saved%20Searches.png)

**Summary Table:** The client wanted a simple way to compare various search terms, so a pivot table was added so that he can easily view the results.

![Summary Table](https://github.com/jbrekes/google_apps_script_projects/blob/main/google_bing_search_results/Summary%20Table.png)

## 📈 Future Development

Future development for this project includes expanding the search engines used and improving the search algorithm to provide more accurate results.

Feel free to clone or fork this repository to use this tool in your own projects!
