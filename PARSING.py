from demoparser2 import DemoParser  # https://github.com/LaihoE/demoparser
import pandas as pd
from openpyxl import load_workbook
import os
from Coloring import highlight_excel_sections # Found online for coloring excel files

def get_df(title: str, ticks):
    # initialize a list to store all round data
    data = []
    # loop through each tick and collect data
    print(title)
    for tick in ticks:
        print(f"tick: {tick}")
        df = parser.parse_ticks(fields, ticks=[tick])
        data.append(df)  # store the dataframe
    # combine all dataframes into a single one
    return pd.concat(data, ignore_index=True)


# PUT DEMO FILE HERE
# PUT DEMO FILE HERE
# PUT DEMO FILE HERE
files = [os.getcwd() + "\\DEMO\\auto-20250403-0032-de_dust2-Counter-Strike_2.dem"]

for file in files:
    parser = DemoParser(file)

    """
    # If you just want the names of all events then you can use this:
    event_names = parser.list_game_events()
    df = parser.parse_events(["all"])
    print(df)
    print(type(df))
    """

    header = parser.parse_header()
    print(f"\naddons: {header["addons"]}")
    print(f"version: {header["demo_version_name"]}")
    print(f"guid: {header["demo_version_guid"]}")
    print(f"server: {header["server_name"]}")
    print(f"client: {header["client_name"]}")
    print(f"map: {header["map_name"]}\n")

    # selecting data to collect
    #player_connect_ticks = parser.parse_event("player_connect")["tick"].tolist()
    #player_activate_ticks = parser.parse_event("player_activate")["tick"].tolist()
    #begin_new_match_ticks = parser.parse_event("begin_new_match")["tick"].tolist()
    round_start_ticks = parser.parse_event("round_start")["tick"].tolist()
    round_freeze_end_ticks = parser.parse_event("round_freeze_end")["tick"].tolist()
    round_end_ticks = parser.parse_event("round_end")["tick"].tolist()
    fields = ["start_balance",
              "balance",
              "cash_spent_this_round",
              "total_cash_spent",
              "is_alive",
              "health",
              "armor_value",
              "life_state",
              "round_start_equip_value",
              "current_equip_value",
              "kills_total",
              "headshot_kills_total",
              "kill_reward_total",
              "alive_time_total",
              "objective_total",
              "mvps",
              "utility_damage_total",
              "3k_rounds_total",
              "4k_rounds_total",
              "ace_rounds_total",
              "enemies_flashed_total",
              "total_rounds_played",
              "team_rounds_total",
              "rounds_played_this_phase",
              "team_name",
              "score",
              "money_saved_total",
              "cash_earned_total",
              "is_connected",
              "ping",
              "game_phase",
              "active_weapon_name",
              "active_weapon_original_owner",
              "weapon_purchases_this_match",
              "weapon_purchases_this_round",
              "has_helmet",
              "armor_value",
              "has_defuser",
              "prev_owner",
              "active_weapon_ammo",
              "total_ammo_left",
              "item_def_idx",
              "item_id_high",
              "item_id_low",
              "inventory_position",
              "is_silencer_on",
              "is_scoped",
              "shots_fired",
              "team_surrendered",
              "is_warmup_period",
              "is_freeze_period",
              "is_terrorist_timeout",
              "is_ct_timeout",
              "round_win_status",
              "round_win_reason",
              "team_score_first_half",
              "team_score_second_half",
              "inventory",
              "inventory_as_ids"]

    # sets up spreadsheet and puts data in spreadsheet
    #player_connect_df = get_df("PLAYER CONNECT:", player_connect_ticks)
    #player_activate_df = get_df("PLAYER ACTIVATE:", player_activate_ticks)
    #begin_new_match_df = get_df("BEGIN NEW:", begin_new_match_ticks)
    round_start_df = get_df("ROUND START:", round_start_ticks)
    round_freeze_end_df = get_df("ROUND FREEZE:", round_freeze_end_ticks)
    round_end_df = get_df("ROUND END:", round_end_ticks)

    # export to excel
    excel_file = os.path.basename(file) + ".xlsx"
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        #player_activate_df.to_excel(writer, sheet_name="Player Activate", index=False)
        #player_connect_df.to_excel(writer, sheet_name="Player Connect", index=False)
        #begin_new_match_df.to_excel(writer, sheet_name="Begin New", index=False)
        round_start_df.to_excel(writer, sheet_name="Round Start", index=False)
        round_freeze_end_df.to_excel(writer, sheet_name="Round Freeze", index=False)
        round_end_df.to_excel(writer, sheet_name="Round End", index=False)

    # reloads file
    wb = load_workbook(excel_file)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for col in ws.columns:
            max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

    # colors sections
    highlight_excel_sections(wb)

    # saves file
    wb.save(excel_file)
    wb.close()