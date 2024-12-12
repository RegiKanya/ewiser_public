import unittest
from datetime import datetime
from gsheet_data_filler import (create_timestamp, create_breakdown_values, filtered_adjusted_inverters, 
                                format_power_diff_ratio, format_power_diff_quantity)

class TestPowerPlantFunctions(unittest.TestCase):

    def test_create_timestamp(self):
        power_plants = [
            {
                'powerPlantId': 1,
                'inverters': [
                    {'startTimestamp': '2023-12-10T15:00:00Z'},
                    {'startTimestamp': '2023-12-11T10:00:00Z'}
                ]
            },
            {
                'powerPlantId': 2,
                'inverters': [
                    {'startTimestamp': '2023-12-09T12:00:00Z'}
                ]
            }
        ]

        expected_output = {
            1: '2023-12-11',
            2: '2023-12-09'
        }

        self.assertEqual(create_timestamp(power_plants), expected_output)

    def test_create_breakdown_values(self):
        power_plants = [
            {
                'powerPlantId': 1,
                'breakDown': {'interval': '2023-12-10T00:00:00Z/2023-12-11T00:00:00Z'}
            },
            {
                'powerPlantId': 2,
                'breakDown': {'interval': '2023-12-09T00:00:00Z/2023-12-09T23:59:59Z'}
            },
            {
                'powerPlantId': 3}
        ]

        expected_output = {
            1: '2023-12-10 -- 2023-12-11',
            2: '2023-12-09 -- 2023-12-09',
            3: 'Nincs'
        }

        self.assertEqual(create_breakdown_values({'body': {'modbus': {'powerPlants': power_plants}}}), expected_output)

    def test_format_power_diff_ratio(self):
        self.assertEqual(format_power_diff_ratio(0.1234), 12.3)
        self.assertEqual(format_power_diff_ratio(None), None)

    def test_format_power_diff_quantity(self):
        self.assertEqual(format_power_diff_quantity(12.3456), 12.3)
        self.assertEqual(format_power_diff_quantity(None), None)

    def test_filtered_adjusted_inverters(self):
        inverters = [
            {
                'powerPlantId': 1,
                'inverterPowerDifference': 15.6789,
                'inverterPowerDifferenceRatio': 0.15,
                'referenceInverterStatus': 'Error'
            },
            {
                'powerPlantId': 2,
                'inverterPowerDifference': 20.1234,
                'inverterPowerDifferenceRatio': 0.05,
                'referenceInverterStatus': 'OK'
            }
        ]
        desired_keys = ['powerPlantId', 'inverterPowerDifference', 'referenceInverterStatus']
        latest_timestamps = {1: '2023-12-11'}
        pp_id_filter = [2]
        break_down_filtered_date = {1: '2023-12-10 -- 2023-12-11'}

        expected_output = [
            [1, None, 'Error', 15.7, 15.0, '2023-12-10 -- 2023-12-11', '2023-12-11']
        ]

        result = filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps, pp_id_filter, break_down_filtered_date)
        self.assertEqual(result, expected_output)

if __name__ == '__main__':
    unittest.main()
