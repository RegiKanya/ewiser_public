# ewiser_public

## Process

## 1. Download the new inverter errors json
   - open the ewiser dashboard
   - Vezérlőpult --> Inverter hibák
   - F12 or right click - Inspect
   - Network and dowload the {;}latest-error-log?siId=424 reponse
   - json has to be named: yyyy-mm-dd.json (today date)
   - go to the script run create_connection.py code
   
## 2. Run Hibák frissítése button
   
These parameters filter out the inverter issue:
'referenceInverterStatus' != OK &
'inverterPowerDifferenceRatio' >= 0.1
('errorInverterCount') > 0
+ hardcoded list about power_plants from which does not need to collect data

## 3. Handle issue and fill up the ACTION sheet with information
## 4. Push the Folyamatban button to move the process forward
## 5. Repeat the process time to time
   - do the manual downloading then run the script every morning
     

(Further Improvements) import the data into the spreadsheet (RAW sheet)
   - using python code to upload data (create the connection, requires: GCP project and enabling Google Sheets API + Google Drive API and service account or OAuth 2.0 Client IDs setup)
   - connect the spreadshet directly with the webpage (e.g.: https://hasdata.com/blog/google-sheets-web-scraping)
