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
#print(latest_timestamps)

# Define the keys you are interested in
desired_keys = [
    'powerPlantId', 'name', 'locationCity', 'locationParcelNumber', 
    'totalInverterCount', 'errorInverterCount', 'breakDown', 
    'referenceInverterStatus', 'inverterPowerDifference', 'inverterPowerDifferenceRatio']
inverters = data['body']['modbus']['powerPlants']

def filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps):
    filtered_inverters = []

    #Collect inverters that meet the criteria and add the latest timestamp to icident date
    for inverter in inverters:
        if all(key in inverter for key in desired_keys) and inverter['referenceInverterStatus'] != "OK" and \
            (inverter['inverterPowerDifference'] is not None and inverter['inverterPowerDifference'] >= 10) and \
                (inverter['inverterPowerDifferenceRatio'] is not None and inverter['inverterPowerDifferenceRatio'] >= 0.4):
                    inverter_data = [inverter[key] for key in desired_keys]
                    power_plant_id = inverter_data[0]

                    if power_plant_id in latest_timestamps:
                         inverter_data.append(latest_timestamps[power_plant_id])

                    filtered_inverters.append(inverter_data)
    return filtered_inverters

result = filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps)

#for item in result: 
#    print(item)

# filtered_inverters now contains the lists of values to be inserted into the Google sheet columns
# Define the scope and the credentials file
SPREADSHEET_ID = 
RANGE_NAME = 

# Update the Google Sheet with the filtered inverters data
update_sheet(SPREADSHEET_ID, RANGE_NAME, filtered_adjusted_inverters)
