import json
from datetime import datetime
import glob

# read the json file in a dynamic way
folder_path = '/Users/kanyaregina/json_files/'
today = datetime.today().strftime('%Y-%m-%d')
json_file = glob.glob(f"{folder_path}{today}.json")
for file_path in json_file:
    with open(file_path, 'r') as f:
        data = json.load(f)

        # get the 'modbus' data
        head = data.get('head', {})
        modbus = head.get('body', {}).get('modbus', {})
        
        # if 'powerPlants' object has inverters, then get it
        power_plants = modbus.get('powerPlants', [])
        
        for plant in power_plants:
            inverters = plant.get('inverters', [])

#create a function which look for timestamp whihch can create the value for 'incident_timestamp'
def create_timestamp(power_plants):
    # Create a dictionary to store the latest timestamps for each power plant
    power_plants_dict = {}

    for power_plant in power_plants:
        power_plant_id = power_plant.get('powerPlantId')
        if power_plant_id not in power_plants_dict:
            power_plants_dict[power_plant_id] = None

        inverters = power_plant.get('inverters', [])
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

# Extract power plants data safely using .get()
power_plants = data.get('body', {}).get('modbus', {}).get('powerPlants', [])
latest_timestamps = create_timestamp(power_plants)

# Define the keys you are interested in
desired_keys = [
    'powerPlantId', 'name', 'locationCity', 'locationParcelNumber', 
    'totalInverterCount', 'errorInverterCount', 'breakDown', 
    'referenceInverterStatus', 'inverterPowerDifference', 'inverterPowerDifferenceRatio']
inverters = data.get('body', {}).get('modbus', {}).get('powerPlants', [])

def filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps):
    filtered_inverters = []

    #Collect inverters that meet the criteria and add the latest timestamp to icident date
    for inverter in inverters:
        # Check for the most restrictive conditions first for performance
        if inverter.get('referenceInverterStatus') != "OK" and \
            (inverter.get('inverterPowerDifference') is not None and inverter['inverterPowerDifference'] >= 10) and \
            (inverter.get('inverterPowerDifferenceRatio') is not None and inverter['inverterPowerDifferenceRatio'] >= 0.4):
                    inverter_data = [inverter[key] for key in desired_keys]
                    power_plant_id = inverter_data[0]

                    if power_plant_id in latest_timestamps:
                         inverter_data.append(latest_timestamps[power_plant_id])

                    filtered_inverters.append(inverter_data)
    return filtered_inverters

result = filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps)

#for item in result: 
#    print(item)

