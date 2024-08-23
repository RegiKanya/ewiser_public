# ewiser_public

## Process
1. import the data into the spreadsheet (RAW sheet)
   - using python code to upload data (create the connection, requires: GCP project and enabling Google Sheets API + Google Drive API and service account or OAuth 2.0 Client IDs setup)
   - connect the spreadshet directly with the webpage (e.g.: https://hasdata.com/blog/google-sheets-web-scraping)
2. Run Hibák frissítése button
These parameters filter out the inverter issue:
'referenceInverterStatus' != OK &
'inverterPowerDifference' is not None and 'inverterPowerDifference' >= 10 &
'inverterPowerDifferenceRatio' is not None and 'inverterPowerDifferenceRatio' >= 0.4
4. Handle issue and fill up the ACTION sheet with information
5. Push the Folyamatban button to move the process forward
6. Repeat the process time to time
   - reqiures that the RAW sheet has to refresh periodically too
   - time-trigger on this function

