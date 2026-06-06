# DecodeLabs Assignment Viewer

This project is a local web app for viewing two Excel assignments in a professional browser interface.

## Files Used

- `DecodeLabs Assignment 1.xlsx`
- `DecodeLabs Assignment 2.xlsx`

The app reads these files from the project folder. If they are not found there, it also checks your Desktop as a fallback.

## How To Run

Double-click:

```text
start_app.bat
```

Keep the terminal window open. The app will print a local URL such as:

```text
http://127.0.0.1:8000
```

Open that URL in your browser.

## Features

- First popup lets the user choose Assignment 1 or Assignment 2
- Professional dashboard-style layout
- Excel sheet tabs
- Search within the selected sheet
- Pagination for large sheets
- Workbook summary cards
- Responsive layout for desktop and mobile screens

## Technology

- Python for reading Excel files and running the local server
- HTML, CSS, and JavaScript for the user interface
- `openpyxl` for Excel workbook support

If `openpyxl` is missing, install it with:

```text
pip install openpyxl
```
