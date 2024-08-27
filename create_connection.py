#create connection with the sheet 
# Importing required library 
import pygsheets 
import pandas as pd
from gsheet_data_filler import filtered_adjusted_inverters,inverters,desired_keys,latest_timestamps

# Create the Client 
client = pygsheets.authorize(service_account_file="...") 
#print(client.spreadsheet_titles()) 

def main():
    sheet = client.open_by_key('...')
    data = filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps)
    #check that it's a dataframe if not then convert
    if isinstance(data, list):
        data = pd.DataFrame(data)
    worksheet = sheet.worksheet_by_title('...')  # or worksheet = sheet.worksheet_by_title('Munkalap neve')
    worksheet.clear()
    worksheet.set_dataframe(data, (1, 1), copy_head=False) #(2nd row, 1st column)
    print("Data has been successfully cleared and reloaded.")

if __name__ == '__main__':
    main()




