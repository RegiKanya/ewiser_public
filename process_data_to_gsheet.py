import json
from datetime import datetime
import glob
import os

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

    for power_plant in power_plants:
        power_plant_id = power_plant.get('powerPlantId')
        if power_plant_id not in break_down_dict:
            break_down_dict[power_plant_id] = 'Nincs'
        
        breakdown = power_plant.get('breakDown', {})
        if breakdown is not None:
            formatted_date = breakdown.get('interval')
            if formatted_date:
                try:
                    start_iso, end_iso = formatted_date.split("/")
        
                    start_dt = datetime.fromisoformat(start_iso.replace('Z', '+00:00'))
                    end_dt = datetime.fromisoformat(end_iso.replace('Z', '+00:00'))
        
                    start_str = start_dt.strftime("%Y-%m-%d")
                    end_str = end_dt.strftime("%Y-%m-%d")

                    break_down_dict[power_plant_id] = f"{start_str} -- {end_str}"
                except ValueError as e:
                    print(f"Error while handling the dates in powerPlants {power_plant_id}: {e}")
                    break_down_dict[power_plant_id] = "Incorret date"
        else:
            break_down_dict[power_plant_id] = 'Nincs'
    
    return break_down_dict

def format_power_diff_ratio(power_diff_ratio):
    if power_diff_ratio is None:
        return None
    return round(power_diff_ratio * 100, 1)

def format_power_diff_quantity(power_diff_quantity):
    if power_diff_quantity is None:
        return None
    return round(power_diff_quantity, 1)

def filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps, pp_id_filter, break_down_filtered_date, is_whitelist=False):
    filtered_inverters = []

    for inverter in inverters:
        power_plant_id = inverter.get('powerPlantId')
        power_diff_ratio = inverter.get('inverterPowerDifferenceRatio') #modified 0.4 --> 0.1
        power_diff_quantity = inverter.get('inverterPowerDifference')
        id_in_list = power_plant_id in pp_id_filter
        id_matches_filter_logic = (is_whitelist and id_in_list) or (not is_whitelist and not id_in_list)

        if id_matches_filter_logic:
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

def process_json_files(file_paths, id_filter, is_whitelist_mode):

    processed_data = []
    for file_path in file_paths:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except FileNotFoundError:
            print(f"🚨 Error: File not found at {file_path}")
            continue
        except json.JSONDecodeError:
            print(f"🚨 Error reading JSON file: {file_path}")
            continue

        power_plants = data.get('body', {}).get('modbus', {}).get('powerPlants', [])


        if not power_plants:
            print(f"🟡 No 'powerPlants' data found in {file_path}")
            continue

        break_down_filtered_date = create_breakdown_values(power_plants)
        latest_timestamps = create_timestamp(power_plants)
        desired_keys = [
            'powerPlantId', 'name', 'locationCity', 'locationParcelNumber',
            'totalInverterCount', 'errorInverterCount', 'referenceInverterStatus']

        result = filtered_adjusted_inverters(power_plants, desired_keys, latest_timestamps, id_filter, break_down_filtered_date, is_whitelist=is_whitelist_mode)
        processed_data.extend(result)

    return processed_data

sources = [
    {
        "name": "MARKET_424",
        "folder": "/Users/kanyaregina/Documents/Ewiser/inverter-errors/json_files/",
        "filter": [144, 145, 146, 216, 381, 382, 383, 408, 409, 415, 455, 712, 713, 1331, 1476, 1501, 1502, 1513, 1516, 1517, 1518, 1546, 1548, 1598, 1620, 1621, 1648, 1649, 1651, 1653, 1655, 1673],
        "is_whitelist": False 
    },
    {
        "name": "MARKET_119",
        "folder": "/Users/kanyaregina/Documents/Ewiser/inverter-errors/market_119/",
        "filter": [1419, 1442, 1443, 1444],
        "is_whitelist": True 
    }
]

today_str = datetime.today().strftime('%Y-%m-%d')
all_results = []

print(f"Processing data for: {today_str}\n")

for source in sources:
    print(f"--- Starting source: {source['name']} ---")

    file_pattern = os.path.join(source['folder'], f"{today_str}.json")
    found_files = glob.glob(file_pattern)

    if not found_files:
        print(f"🟡 No JSON file found for today in folder: '{source['folder']}'\n")
        continue 

    print(f"Found file(s): {found_files}")

    result_from_source = process_json_files(
        file_paths=found_files,
        id_filter=source['filter'],
        is_whitelist_mode=source['is_whitelist']
    )

    print(f"Processed {len(result_from_source)} items from this source.")
    all_results.extend(result_from_source)

print(f"✅ Processing complete! The final combined list contains {len(all_results)} items.")

# Optional: Print the first few items to verify
#print("\nFirst 30 items in the combined list:")
#for item in all_results[:30]:
#    print(item)