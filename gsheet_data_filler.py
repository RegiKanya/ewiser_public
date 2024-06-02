import json
from datetime import datetime
from create_connection import update_sheet #import the function from create_connection.py file
import glob

# read the json file in a dynamic way
folder_path = #insert your local folder path
json_file = glob.glob(folder_path + '*.json')
for file_path in json_file:
    with open(file_path, 'r') as f:
        data = json.load(f)

#create a function which look for timestamp whihch can create the value for 'incident_timestamp'
def create_timestamp(power_plants):
    # Create a dictionary to store the latest timestamps for each power plant
    power_plants_dict = {}

    for power_plant in power_plants:
        power_plant_id = power_plant['powerPlantId']
        if power_plant_id not in power_plants_dict:
            power_plants_dict[power_plant_id] = None

        inverters = power_plant['inverters']
        for inverter in inverters:
            timestamp_str = inverter.get('startTimestamp')
            if timestamp_str:
                # Convert the timestamp to a date object
                timestamp_date = datetime.fromisoformat(timestamp_str.rstrip('Z'))
                formatted_date = timestamp_date.strftime('%Y-%m-%d')

                # Update the dictionary if this timestamp is newer
                if not power_plants_dict[power_plant_id] or formatted_date > power_plants_dict[power_plant_id]:
                    power_plants_dict[power_plant_id] = formatted_date

    return power_plants_dict

# Extract power plants data
power_plants = data['body']['modbus']['powerPlants']
latest_timestamps = create_timestamp(power_plants)

# Define the keys you are interested in
desired_keys = [
    'powerPlantId', 'name', 'locationCity', 'locationParcelNumber', 
    'totalInverterCount', 'errorInverterCount', 'breakDown', 
    'referenceInverterStatus', 'inverterPowerDifference', 'inverterPowerDifferenceRatio']

# Collect inverters that meet the criteria
inverters = data['body']['modbus']['powerPlants']
filtered_inverters = []

for inverter in inverters:
    if all(key in inverter for key in desired_keys) and inverter['referenceInverterStatus'] != "OK" and (inverter['inverterPowerDifference'] is not None and inverter['inverterPowerDifference'] >= 10) and (inverter['inverterPowerDifferenceRatio'] is not None and inverter['inverterPowerDifferenceRatio'] >= 0.4):
        filtered_inverters.append([inverter[key] for key in desired_keys])

for inverter in filtered_inverters:
    print(f'these are the issues: {inverter}')

# filtered_inverters now contains the lists of values to be inserted into the Google sheet columns
# Define the scope and the credentials file
SPREADSHEET_ID = 
RANGE_NAME = 

# Update the Google Sheet with the filtered inverters data
update_sheet(SPREADSHEET_ID, RANGE_NAME, filtered_inverters)
