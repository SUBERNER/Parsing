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

    ticks.sort()  # sorts the ticks in the list

    # displays all ticks information
    print(f"{len(ticks)} total ticks: {ticks}\n")

    # parses data from each tick
    all_data = []  # stores all data form parsed ticks from all events
    start_data = []  # stores all data from parsed ticks form stating a round
    freeze_end_data = []  # stores all data from parsed ticks form the end of a buying phase of a round
    end_data = []  # stores all data from parsed ticks form the end of a round
    for tick in ticks:
        print(f"tick: {tick}")
        df = parser.parse_ticks(fields, ticks=[tick])
        # separates and categorizes all the ticks into the data lists
        if tick in start_ticks:
            start_data.append(df)
        elif tick in freeze_end_ticks:
            freeze_end_data.append(df)
        elif tick in end_ticks:
            end_data.append(df)
        all_data.append(df)  # stores all events in one location

    # combines data from all ticks together and converts to csv
    csv_name = os.path.splitext(os.path.basename(file))[0] + ".csv"  # name of csv file

    # creates and combines data from rows into a row for each player wirth all the data for each round
    demo_columns = ['user_id',
                    'total_rounds_played',
                    'start_inventory_as_ids',
                    'end_inventory_as_ids',
                    'start_balance',
                    'total_cash_spent',
                    'cash_spent_this_round',
                    'kills_total',
                    'headshot_kills_total',
                    'damage_total',
                    'utility_damage_total',
                    'enemies_flashed_total',
                    'objective_total',
                    'deaths_total',
                    'alive_time_total',
                    'life_state',
                    'cash_earned_total',
                    'kill_reward_total',
                    'equipment_value_total',
                    'current_equip_value',
                    'round_start_equip_value',
                    'mvps']  # names of new columns

    # creates a new empty csv file
    demo_df = pd.DataFrame(columns=demo_columns)


    rows = []

    # iterate over each row of data from each dataset reorganizing them
    for start_df, freeze_df, end_df in zip(start_data, freeze_end_data, end_data):
        for index, end_row in end_df.iterrows():  # iterates over each row

            rows.append({'user_id': end_row['user_id'],
                         'total_rounds_played': end_row['total_rounds_played'],
                         'start_inventory_as_ids': end_row['inventory_as_ids'],  # WILL CHANGE IN THE FUTURE
                         'end_inventory_as_ids': end_row['inventory_as_ids'],
                         'start_balance': end_row['start_balance'],
                         'total_cash_spent': end_row['total_cash_spent'],
                         'cash_spent_this_round': end_row['cash_spent_this_round'],
                         'kills_total': end_row['kills_total'],
                         'headshot_kills_total': end_row['headshot_kills_total'],
                         'damage_total': end_row['damage_total'],
                         'utility_damage_total': end_row['utility_damage_total'],
                         'enemies_flashed_total': end_row['enemies_flashed_total'],
                         'objective_total': end_row['objective_total'],
                         'deaths_total': end_row['deaths_total'],
                         'alive_time_total': end_row['alive_time_total'],
                         'life_state': end_row['life_state'],
                         'cash_earned_total': end_row['cash_earned_total'],
                         'kill_reward_total': end_row['kill_reward_total'],
                         'equipment_value_total': end_row['equipment_value_total'],
                         'current_equip_value': end_row['current_equip_value'],
                         'round_start_equip_value': end_row['round_start_equip_value'],
                         'mvps': end_row['mvps']})

    # adds all the rows at the end
    demo_df = pd.DataFrame(rows)


    #demo_df = pd.concat(end_data + start_data, ignore_index=True)
    #demo_df.sort_values(by='tick', inplace=True)
    #demo_df.drop(columns=['steamid', 'name'], inplace=True)

    # edits the csv files
    demo_df.to_csv(csv_name, index=False)

