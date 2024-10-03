import unittest
from datetime import datetime
from unittest.mock import patch
from gsheet_data_filler import create_timestamp, create_breakdown_values, filtered_adjusted_inverters 

class TestDataFunctions(unittest.TestCase):

    def test_create_timestamp(self):
        # Sample data for testing
        power_plants = [
            {
                'powerPlantId': 1,
                'inverters': [
                    {'startTimestamp': '2023-10-26T10:00:00Z'},
                    {'startTimestamp': '2023-10-27T12:00:00Z'} 
                ]
            },
            {
                'powerPlantId': 2,
                'inverters': [
                    {'startTimestamp': '2023-10-26T08:00:00Z'}
                ]
            }
        ]

        expected_timestamps = {
            1: '2023-10-27',
            2: '2023-10-26'
        }
        self.assertEqual(create_timestamp(power_plants), expected_timestamps)

    def test_create_timestamp_empty_inverters(self):
        power_plants = [
            {'powerPlantId': 1, 'inverters': []}
        ]
        expected_timestamps = {1: None}
        self.assertEqual(create_timestamp(power_plants), expected_timestamps)

    def test_create_breakdown_values(self):
        power_plants = [
            {
                'powerPlantId': 1,
                'inverters': [
                    {'breakDown': {'interval': '2023-10-25T10:00:00Z/2023-10-27T12:00:00Z'}}
                ]
            },
            {
                'powerPlantId': 2,
                'inverters': [
                    {'breakDown': {'interval': None}}
                ]
            }
        ]
        expected_breakdown_values = {
            1: '2023-10-25 - 2023-10-27',
            2: 'Nincs'
        }
        self.assertEqual(create_breakdown_values(power_plants), expected_breakdown_values)

    def test_create_breakdown_values_no_breakdown(self):
        power_plants = [
            {
                'powerPlantId': 1,
                'inverters': [
                    {} 
                ]
            }
        ]
        expected_breakdown_values = {
            1: 'Nincs'
        }
        self.assertEqual(create_breakdown_values(power_plants), expected_breakdown_values)

    def test_filtered_adjusted_inverters(self):
        inverters = [
            # Matches criteria
            {'powerPlantId': 1, 'name': 'Test Inverter 1', 'locationCity': 'City A', 'locationParcelNumber': '1234', 
             'totalInverterCount': 10, 'errorInverterCount': 2, 'referenceInverterStatus': 'ERROR', 
             'inverterPowerDifference': 1000, 'inverterPowerDifferenceRatio': 0.15},
            # Doesn't match - powerPlantId in inverters_id_filter
            {'powerPlantId': 408, 'name': 'Test Inverter 2', 'locationCity': 'City B', 'locationParcelNumber': '5678', 
             'totalInverterCount': 5, 'errorInverterCount': 1, 'referenceInverterStatus': 'ERROR', 
             'inverterPowerDifference': 500, 'inverterPowerDifferenceRatio': 0.2},
            # Doesn't match - referenceInverterStatus is OK
            {'powerPlantId': 2, 'name': 'Test Inverter 3', 'locationCity': 'City C', 'locationParcelNumber': '9012', 
             'totalInverterCount': 8, 'errorInverterCount': 0, 'referenceInverterStatus': 'OK', 
             'inverterPowerDifference': 800, 'inverterPowerDifferenceRatio': 0.1},
            # Doesn't match - inverterPowerDifferenceRatio is too low
            {'powerPlantId': 3, 'name': 'Test Inverter 4', 'locationCity': 'City D', 'locationParcelNumber': '3456', 
             'totalInverterCount': 12, 'errorInverterCount': 3, 'referenceInverterStatus': 'ERROR', 
             'inverterPowerDifference': 1200, 'inverterPowerDifferenceRatio': 0.05}
        ]
        desired_keys = ['powerPlantId', 'name', 'locationCity', 'locationParcelNumber', 
                        'totalInverterCount', 'errorInverterCount', 'inverterPowerDifference', 'inverterPowerDifferenceRatio']
        latest_timestamps = {1: '2023-10-28', 2: '2023-10-27'}
        inverters_id_filter = [408, 409]
        break_down_filtered_date = {1: '2023-10-26 - 2023-10-28', 2: 'Nincs'}

        expected_filtered_inverters = [
            [1, 'Test Inverter 1', 'City A', '1234', 10, 2, '2023-10-26 - 2023-10-28', 1000, 0.15, '2023-10-28']
        ]
        
        self.assertEqual(filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps, inverters_id_filter, break_down_filtered_date), 
                         expected_filtered_inverters)

if __name__ == '__main__':
    unittest.main()
