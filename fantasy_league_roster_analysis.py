import logging

from fantasy_league_stats_manual_analysis import reg_season_weeks

logging.basicConfig(level=logging.DEBUG, format='%(levelname)s: %(message)s')
import os
import re
import pandas as pd

import fantasy_league_stats_manual_analysis as fls

roster_csvs_folder_path = r"C:\Users\aaron\PycharmProjects\oct25_fantasy_stats_with_scraping\output_csvs\2457715_history_teamgamecenter"
players = fls.players
league_set_up = fls.league_set_up
typical_columns = pd.read_csv(os.path.join(roster_csvs_folder_path, "2014", "3.csv")).columns
# typically will have all 10 normal roster spots with 5 bench spots. Thus, normally, 37 columns.
# If more bench players than normal, then number of total columns for csv for that week = 37 + (excess bench spots * 2)

# the below dictionary contains the associated df for each week so I don't have to go looking
yearly_points_bench_lineup_dict = {year:[0,0] for year in range(2014,2025)}
df_dict_for_weeks = {}
for year in range(2014,2025):
    current_year_csv_path = os.path.join(roster_csvs_folder_path, str(year))
    for week in range(1, (league_set_up[year]["reg_season_weeks"] + 2)):
        # adding current df to df_dict
        current_week_csv = os.path.join(current_year_csv_path, f"{str(week)}.csv")
        current_week_df = pd.read_csv(current_week_csv, index_col=0)
        df_dict_for_weeks[f"{year}_week{week}"] = current_week_df

        for owner in current_week_df.index:
            #because both kevins have same owner name in dataset
            # first what to do if both kevins are playing each other in a given week
            if owner == "Kevin" and current_week_df["Kevin.1"]["Opponent"] == "Kevin":
                # fix later
                print(f"Kevins play each other {year}, week{week}")

            elif owner == "Kevin":
                #check opponent dict from manual analysis.




# which statistics to calculate, first: total points for week in lineup vs bench,
# then how common each player for each manager


