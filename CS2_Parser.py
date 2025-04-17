from demoparser2 import DemoParser  # https://github.com/LaihoE/demoparser
import pandas as pd
from openpyxl import load_workbook
import os

# stores all the demo files that will be scanned and parsed
dem_files = [os.getcwd() + "\\DEMO\\better_TEST.dem"]
# stores all fields that will be extracted and parsed into the csv file
fields = ['inventory_as_ids', 'life_state', 'kills_total',
          'headshot_kills_total', 'damage_total', 'deaths_total',
          'objective_total', 'utility_damage_total', 'enemies_flashed_total',
          'equipment_value_total', 'kill_reward_total', 'cash_earned_total',
          'alive_time_total', 'user_id', 'mvps',
          'start_balance', 'total_cash_spent', 'cash_spent_this_round',
          'round_start_equip_value', 'current_equip_value', 'total_rounds_played']

for file in dem_files:
    parser = DemoParser(file)

    # gives basic information about the demo file
    header = parser.parse_header()
    print(f"\nfile: {os.path.basename(file)}")
    print(f"version: {header["demo_version_name"]}")
    print(f"guid: {header["demo_version_guid"]}")
    print(f"server: {header["server_name"]}")
    print(f"client: {header["client_name"]}")
    print(f"map: {header["map_name"]}")

    # gets all events assigned to the end and start of a round, and when players are around to move at the start of the round
    ticks = []  # stores all ticks in one place
    start_ticks = parser.parse_event("round_start")["tick"].tolist()
    print(f"\n{len(start_ticks)} round start ticks: {start_ticks}")
    ticks.extend(start_ticks)

    freeze_end_ticks = parser.parse_event("round_freeze_end")["tick"].tolist()
    print(f"{len(freeze_end_ticks)} round freeze end ticks: {freeze_end_ticks}")
    ticks.extend(freeze_end_ticks)

    end_ticks = parser.parse_event("round_end")["tick"].tolist()
    print(f"{len(end_ticks)} round end ticks: {end_ticks}")
    ticks.extend(end_ticks)

    ticks.sort() # sorts the ticks in the list

    # displays all ticks information
    print(f"{len(ticks)} total ticks: {ticks}\n")

    # parses data from each tick
    data = []  # stores all data from parsed ticks
    for tick in ticks:
        print(f"tick: {tick}")
        df = parser.parse_ticks(fields, ticks=[tick])
        data.append(df)  # store the dataframe

    # combines data from all ticks together and converts to csv
    demo_df = pd.concat(data, ignore_index=True)
    demo_df.drop(columns=['steamid', 'name'], inplace=True)
    demo_df.to_csv('output.csv', index=False)

