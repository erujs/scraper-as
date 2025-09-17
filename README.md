# scraper-as

Table Scraper ~ Google Apps Script project that extracts `<table>` elements from any webpage and pastes them into a Google Sheet.  

## Setup Instructions

1. Open any Google Sheet (create a new one if needed).
2. Go to Extensions > App Script.
3. Delete any code in `Code.gs` and paste the contents of `Code.gs` from this repo.
4. Save the code. You can simply press `Ctrl+S` to save.

## How to Run

1. In the Apps Script editor, select `scrapeTables` from the function dropdown, then click `Run` to execute it.
2. The first time you run it, Google will ask for authorization. Accept the permissions.
3. Once finished, the extracted tables will be inserted into Sheet1. Make sure to read the disclaimer in the code.

## Configuration

- In `Code.gs` in Apps Scripts, locate the following line and update it:

```javascript
// 👇 Change this URL to the page you want to scrape
const url = "https://www.w3schools.com/html/html_tables.asp";
```

- You can also modify the sheet name where tables are inserted:

```javascript
// 👇 Modify the sheet name below to your preferred sheet. If it does not exist, it will be created.
const sheet = ss.getSheetByName("Sheet1") || ss.insertSheet("Sheet1");
```

## Output

- Tables will be written into the sheet specified in the code.
- Each table is prefixed with a label, e.g. Table 1, Table 2, etc.
- A blank line is inserted between tables.

## Limitations

- The script only extracts standard `<table>` elements from the raw HTML.
- It does not execute JavaScript on the target page. Tables generated dynamically by JS will not be captured.
- Works best with static HTML tables.

## Notes

- This project is an upgrade from the previous `vba-webdata` repo. Since VBA is approaching the end of its mainstream support and requires Excel for execution.
- This Google Apps Script version provides a more stable, modern, and free solution for extracting table data from webpages.
- For exporting the scraped data to a new Google Sheet or Excel file, use the companion project [export-as](https://github.com/erujs/export-as).  
These two projects work well together.
- This project is designed to be simple, lightweight, and entirely free to use with Google Sheets.
- We retain `Module.bas` in the repository for anyone who wishes to continue using the legacy VBA version.

## 🧾 License
MIT — do whatever you want with it.

✨ Happy coding!
If you find this project useful, a ⭐ on the repo is always appreciated!