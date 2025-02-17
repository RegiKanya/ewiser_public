import pygsheets 
import pandas as pd
from datetime import datetime
from gsheet_data_filler import filtered_adjusted_inverters, power_plants, desired_keys,latest_timestamps, json_file, pp_id_filter, break_down_filtered_date

# Create the Client 
client = pygsheets.authorize(service_account_file="/Users/kanyaregina/Documents/Ewiser/inverter-errors/access_details/gsheet_ewiser.json") 
#print(client.spreadsheet_titles()) 

def main():
    sheet = client.open_by_key('1rv3EPJ8OLlZq2foVWexufG2THvwU_IFXH-GdVlWK938')
    data = filtered_adjusted_inverters(power_plants, desired_keys, latest_timestamps, pp_id_filter, break_down_filtered_date)
    date = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    #check that it's a dataframe if not then convert
    if isinstance(data, list):
        data = pd.DataFrame(data)
    worksheet = sheet.worksheet_by_title('RAW')  # or worksheet = sheet.worksheet_by_title('Munkalap neve')
    worksheet.clear()
    worksheet.update_value('A1',f"Utoljára frissítve: {date}")
    worksheet.set_dataframe(data, (2, 1), copy_head=False) #(2nd row, 1st column)
    print(f"✅ DONE - Data has been successfully cleared and reloaded {json_file}.")

if __name__ == '__main__':
    main()
 