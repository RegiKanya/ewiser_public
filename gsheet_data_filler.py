import json
from datetime import datetime
import glob

#create a function which look for timestamp whihch can create the value for 'incident_timestamp'
def create_timestamp(power_plants):
    power_plants_dict = {}

    for power_plant in power_plants:
        power_plant_id = power_plant.get('powerPlantId')
        if power_plant_id not in power_plants_dict:
            power_plants_dict[power_plant_id] = None

        inverters = power_plant.get('inverters', [])
        for inverter in inverters:
            timestamp_str = inverter.get('startTimestamp')
            if timestamp_str:
                timestamp_date = datetime.fromisoformat(timestamp_str.rstrip('Z'))
                formatted_date = timestamp_date.strftime('%Y-%m-%d')

                if not power_plants_dict[power_plant_id] or formatted_date > power_plants_dict[power_plant_id]:
                    power_plants_dict[power_plant_id] = formatted_date

    return power_plants_dict

def create_breakdown_values(power_plants):
    break_down_dict = {}

    power_plants = data.get('body', {}).get('modbus',{}).get('powerPlants', [])

    for power_plant in power_plants:
        power_plant_id = power_plant.get('powerPlantId')
        if power_plant_id not in break_down_dict:
            break_down_dict[power_plant_id] = 'Nincs'
        
        #inverters = power_plant.get('inverters', [])
        breakdown = power_plant.get('breakDown', {})
        if breakdown is not None:
            formatted_date = breakdown.get('interval')
            if formatted_date:
                try:
                    start_iso, end_iso = formatted_date.split("/")
        
                    # Parse the ISO 8601 dates into datetime objects
                    start_dt = datetime.fromisoformat(start_iso.replace('Z', '+00:00'))
                    end_dt = datetime.fromisoformat(end_iso.replace('Z', '+00:00'))
        
                    # Format the dates into "yyyy-MM-dd"
                    start_str = start_dt.strftime("%Y-%m-%d")
                    end_str = end_dt.strftime("%Y-%m-%d")

                    break_down_dict[power_plant_id] = f"{start_str} -- {end_str}"
                except ValueError as e:
                    print(f"Error while handling the dates in powerPlants {power_plant_id}: {e}")
                    break_down_dict[power_plant_id] = "Incorret date"
        else:
            break_down_dict[power_plant_id] = 'Nincs'
    
    # Return after the outer loop finishes processing all power plants
    return break_down_dict


#make the power difference ratio prettier and multiple with 100 to show the percantage
def format_power_diff_ratio(power_diff_ratio):
    if power_diff_ratio is None:
        return None
    return round(power_diff_ratio * 100, 1)

def format_power_diff_quantity(power_diff_quantity):
    if power_diff_quantity is None:
        return None
    return round(power_diff_quantity, 1)


def filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps, pp_id_filter, break_down_filtered_date):
    filtered_inverters = []

    #Collect inverters that meet the criteria and modify breakdown value + add the latest timestamp to icident date
    for inverter in inverters:
        power_plant_id = inverter.get('powerPlantId')
        #power_diff = inverter.get('inverterPowerDifference') #do not need to filter
        power_diff_ratio = inverter.get('inverterPowerDifferenceRatio') #modified 0.4 --> 0.1
        power_diff_quantity = inverter.get('inverterPowerDifference')

        # Check for the conditions
        if  power_plant_id not in pp_id_filter:
            if (inverter.get('referenceInverterStatus') != "OK" or power_diff_ratio is not None and power_diff_ratio >= 0.1
                or inverter.get('errorInverterCount') > 0):
                    inverter_data = [inverter.get(key, None) for key in desired_keys]
                    
                    formatted_power_diff_quantity = format_power_diff_quantity(power_diff_quantity)
                    inverter_data.insert(10, formatted_power_diff_quantity)

                    formatted_power_diff_ratio = format_power_diff_ratio(power_diff_ratio)
                    inverter_data.insert(10, formatted_power_diff_ratio)

                    if power_plant_id in latest_timestamps:
                        inverter_data.append(latest_timestamps[power_plant_id])
                                        
                    if power_plant_id in break_down_filtered_date:
                        inverter_data.insert(6, break_down_filtered_date[power_plant_id])

                    filtered_inverters.append(inverter_data)
        else: 
            pass
    return filtered_inverters

# read the json file in a dynamic way
folder_path = '/Users/kanyaregina/Documents/Ewiser/inverter-errors/json_files/'
today = datetime.today().strftime('%Y-%m-%d')
json_file = glob.glob(f"{folder_path}{today}.json")

if not json_file:
    print(f"🚨 No JSON file found for {today}.")
else:
    for file_path in json_file:
        with open(file_path, 'r') as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                print(f"🚨 Error reading JSON file: {file_path}")
                continue

            modbus_data = data.get('body', {}).get('modbus', {})
            power_plants = modbus_data.get('powerPlants', [])

            if not power_plants:  
                print(f"🚨 No power plants found in {file_path}")
                continue

            break_down_filtered_date = create_breakdown_values(power_plants)

            desired_keys = [
                'powerPlantId', 'name', 'locationCity', 'locationParcelNumber', 
                'totalInverterCount', 'errorInverterCount', 
                'referenceInverterStatus']
            pp_id_filter = [144, 145, 146, 216, 381, 382, 383, 408, 409, 415, 455, 712, 1331, 1476, 
                            1501, 1502, 1513, 1516, 1517, 1518, 1546, 1548, 1598, 1620, 1621, 1648, 
                            1648, 1649, 1649, 1651, 1651, 1653, 1653, 1655, 1673]

            latest_timestamps = create_timestamp(power_plants)  
        
            # Process the inverters and filter them
            result = filtered_adjusted_inverters(power_plants, desired_keys, latest_timestamps, pp_id_filter, break_down_filtered_date)


            #for item in result: 
                #print(item)

