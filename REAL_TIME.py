from demoparser2 import DemoParser  # https://github.com/LaihoE/demoparser
import pandas as pd
from openpyxl import load_workbook
import os
from Coloring import highlight_excel_sections # Found online for coloring excel files

# PUT DEMO FILE HERE
# PUT DEMO FILE HERE
# PUT DEMO FILE HERE
files = [os.getcwd() + "\\DEMO\\auto-20250403-0032-de_dust2-Counter-Strike_2.dem"]

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

for file in files:
    parser = DemoParser(file)

    all_ticks_df = parser.parse_ticks(fields)

    # export to excel
    excel_file = os.path.basename(file) + ".xlsx"
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        # player_activate_df.to_excel(writer, sheet_name="Player Activate", index=False)
        # player_connect_df.to_excel(writer, sheet_name="Player Connect", index=False)
        # begin_new_match_df.to_excel(writer, sheet_name="Begin New", index=False)
        all_ticks_df.to_excel(writer, sheet_name="All", index=False)

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

    print("PARSER DONE")
    print("PARSER DONE")
    print("PARSER DONE")
    print("PARSER DONE")
    print("PARSER DONE")

