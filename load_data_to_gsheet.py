import pygsheets 
import pandas as pd
from datetime import datetime
from process_data_to_gsheet import all_results, today_str

# Create the Client 
client = pygsheets.authorize(service_account_file="/Users/kanyaregina/Documents/Ewiser/inverter-errors/access_details/gsheet_ewiser.json") 

def main():
    sheet = client.open_by_key('1rv3EPJ8OLlZq2foVWexufG2THvwU_IFXH-GdVlWK938')
    data = all_results
    date = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(data, list):
        data = pd.DataFrame(data)
    worksheet = sheet.worksheet_by_title('RAW')  # or worksheet = sheet.worksheet_by_title('Munkalap neve')
    worksheet.clear(start='A2')
    worksheet.update_value('M1',f"Utoljára frissítve: {date}") 
    worksheet.set_dataframe(data, (2, 1), copy_head=False) #(2nd row, 1st column)
    print(f"✅ DONE - Data has been successfully cleared and reloaded for:'{today_str}'.")

if __name__ == '__main__':
    main()
 