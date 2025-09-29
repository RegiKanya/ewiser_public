import unittest
from datetime import datetime
from ewiser_public.process_data_to_gsheet import create_breakdown_values, create_timestamp, filtered_adjusted_inverters

class TestPowerPlantFunctions(unittest.TestCase):

    def test_create_timestamp(self):
        power_plants = [
                     {
                    "powerPlantId": 1624,
                    "name": "Eurotrade Gamma Kft - Szeghalom 0800/32",
                    "locationCity": "Szeghalom",
                    "locationParcelNumber": "0800/32",
                    "dataLoggerTypes": [
                        "HUAWEI"
                    ],
                    "usersToNotify": [],
                    "inverters": [],
                    "totalInverterCount": 9,
                    "errorInverterCount": 0,
                    "unresolvedErrorInverterCount": 0,
                    "status": "GREEN",
                    "powerPlantComment": None,
                    "powerControlGroupId": 2,
                    "breakDown": None, 
                    "referenceInverterStatus": "OK",
                    "inverterPowerDifference": 11.968237294440634,
                    "inverterPowerDifferenceRatio": 0.025573156612052637,
                    "inverterPowerEnergyDifference": 6.4208340000000135
                },
                {
                    "powerPlantId": 1178,
                    "name": "Dendi Sun Kft. Kocsola 030/1",
                    "locationCity": "Kocsola",
                    "locationParcelNumber": "033/17",
                    "dataLoggerTypes": [
                        "FRONIUS"
                    ],
                    "usersToNotify": [],
                    "inverters": [],
                    "totalInverterCount": 19,
                    "errorInverterCount": 0,
                    "unresolvedErrorInverterCount": 0,
                    "status": "GREEN",
                    "powerPlantComment": None,
                    "powerControlGroupId": 2,
                    "breakDown": None,
                    "referenceInverterStatus": "OK",
                    "inverterPowerDifference": 6.054803376867195,
                    "inverterPowerDifferenceRatio": 0.01213387450273987,
                    "inverterPowerEnergyDifference": 3.7331187499998464
                }
        ]
        expected = {
            1: "2023-01-03",
            2: "2023-02-01"
        }
        result = create_timestamp(power_plants)
        self.assertEqual(result, expected)

    def test_create_breakdown_values(self):
        data = [
                        {
                    "powerPlantId": 1607,
                    "name": "Helionergy Oberon Kft. - Sarkad 0836/17-23",
                    "locationCity": "Sarkad",
                    "locationParcelNumber": "0836/17-23",
                    "dataLoggerTypes": [],
                    "usersToNotify": [],
                    "inverters": [],
                    "totalInverterCount": 0,
                    "errorInverterCount": 0,
                    "unresolvedErrorInverterCount": 0,
                    "status": "GREEN",
                    "powerPlantComment": None,
                    "powerControlGroupId": None,
                    "breakDown": {
                        "breakdownEventId": 10243,
                        "powerPlantId": 1607,
                        "breakdownPowerWatt": 4172000,
                        "startDateTime": "2024-01-29T09:00:00.000Z",
                        "endDateTime": "2025-03-01T07:00:00.000Z",
                        "creationMode": "Manual",
                        "editMode": "Manual",
                        "note": "Termelés megkezdése előtti idő",
                        "interval": "2024-01-29T09:00:00.000Z/2025-03-01T07:00:00.000Z"
                    },
                    "referenceInverterStatus": "OK",
                    "inverterPowerDifference": 0.0,
                    "inverterPowerDifferenceRatio": 0.0,
                    "inverterPowerEnergyDifference": 0.0
                },
                {
                    "powerPlantId": 1660,
                    "name": "DPV Beta Kft - Döge 0148/11-12",
                    "locationCity": "Döge",
                    "locationParcelNumber": "0148/11-12",
                    "dataLoggerTypes": [],
                    "usersToNotify": [],
                    "inverters": [],
                    "totalInverterCount": 0,
                    "errorInverterCount": 0,
                    "unresolvedErrorInverterCount": 0,
                    "status": "GREEN",
                    "powerPlantComment": None,
                    "powerControlGroupId": None,
                    "breakDown": None,
                    "referenceInverterStatus": "OK",
                    "inverterPowerDifference": 0.0,
                    "inverterPowerDifferenceRatio": 0.0,
                    "inverterPowerEnergyDifference": 0.0
                }
        ]
        expected = {
            1: "2023-01-01 -- 2023-01-02",
            2: "2023-02-01 -- 2023-02-03",
            3: "Nincs"
        }
        result = create_breakdown_values(data)
        self.assertEqual(result, expected)


    def test_filtered_adjusted_inverters(self):
        inverters = [
            {
                "powerPlantId": 1,
                "inverterPowerDifferenceRatio": 0.15,
                "inverterPowerDifference": 100.0,
                "referenceInverterStatus": "ERROR",
                "name": "Inverter1",
                "locationCity": "CityA"
            },
            {
                "powerPlantId": 2,
                "inverterPowerDifferenceRatio": 0.05,
                "inverterPowerDifference": 50.0,
                "referenceInverterStatus": "OK",
                "name": "Inverter2",
                "locationCity": "CityB"
            }
        ]
        desired_keys = ["powerPlantId", "name", "locationCity"]
        latest_timestamps = {1: "2023-01-03", 2: "2023-02-01"}
        pp_id_filter = []
        break_down_filtered_date = {1: "2023-01-01 -- 2023-01-02"}

        expected = [
            [1, "Inverter1", "CityA", None, 100.0, 15.0, "2023-01-01 -- 2023-01-02", "2023-01-03"]
        ]

        result = filtered_adjusted_inverters(inverters, desired_keys, latest_timestamps, pp_id_filter, break_down_filtered_date)
        self.assertEqual(result, expected)

if __name__ == "__main__":
    unittest.main()