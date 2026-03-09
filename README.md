# Price File App

Python application for importing, validating, and exporting pharmacy price files.

## Features

- Load supplier price files (CSV)
- Validate pricing data
- Export to SQLite / Access-compatible formats

## Requirements
- Suppliers send csv and Excel files to Healthsoft for import into RxOne/\. These Price Files are used by RxOne Pharmacies to order stock from various suppliers
- Suppliers send the file to RxOne, where it is processed and imported into an Access Database called SUppliers.mdb (at a later date this will be changed to MS SQL)
-currently the app is programmed in VB6
- the app needs to be fast and make it easy for the user to import a file, edit and then export to the Access database.
- main rules
    -- tradename or description can be no longer than 40chars. a file called replacements.csv contains most of the rules for shortening. user needs to be able to manually shorten, and add rules to the file
    -- cost price and retail can be $0 (this is not common)
-when the files is loaded, the user should be able to see changes between the access database and the new price file. there should be an option to sort with changes at the top or with invalid descriptions at the top
- user should be able to export the amended data to access, or create a CSV file from this. 


## Setup

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt

## Starting the app
```

streamlit run app.py
