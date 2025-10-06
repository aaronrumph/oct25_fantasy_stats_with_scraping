import re
from collections import Counter
import logging
logging.basicConfig(level=logging.DEBUG, format='%(levelname)s: %(message)s')
import xlsxwriter
import csv
from pickletools import stringnl_noescape
from re import match
import os
import pandas as pd
from io import StringIO
from numpy.f2py.crackfortran import word_pattern
from numpy.ma.extras import average

# important that all functions, etc.
# from fantasy_league_stats... should be prexifixed with fls.
import fantasy_league_stats_manual_analysis as fls
from fantasy_league_stats_manual_analysis import name_mapping, players

# now onto visualizing in Excel:
# Wwat to visualize:
# all_time_record, all_time_score, h2hrecord, h2hscore, times made playoffs (bar chart), playoff record, finished (regular season and final), points for and against, largest victories, closest victories, largest and smallest scores, largest scores by player


# create workbook #

workbook = xlsxwriter.Workbook("basic_stats_output.xlsx")
# formatting #
fmt_black = workbook.add_format({'bg_color': '#000000', 'font_color': '#000000'})
fmt_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
fmt_bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
fmt_normal = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter'})
fmt_highlight = workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000', 'bold': False, 'align': 'center', 'valign': 'vcenter'})
# making it so displays presentable names instead of database names (Kevin R instead of Rockmael) #
sorted_players = sorted(fls.players)
mapped_names = [fls.name_mapping[player] for player in sorted_players]
name_mapping =fls.name_mapping
df1 =fls.fantasy_df


# ws1: overall record #
logging.debug("Creating worksheet 1")

yearly_records = fls.calculate_reg_records(df1)[2]

# formatting
ws1 = workbook.add_worksheet("Regular Season Record")
ws1.set_column_pixels(0, 0, 90, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws1.set_column_pixels(i, i, 90, fmt_center)
for i in range(12):  # 11 years + header
    ws1.set_row_pixels(i, 24, fmt_center)

# write headers
ws1.write(0, 0, "Year", fmt_bold)
for i, name in enumerate(mapped_names):
    ws1.write(0, i + 1, name, fmt_bold)

# write year labels and records
for i, year in enumerate(range(2014, 2025)):
    ws1.write(i + 1, 0, year, fmt_bold)

    # write each player's record for that year
    for j, player in enumerate(sorted_players):
        if year in yearly_records[player]:
            record = yearly_records[player][year]
            record_str = f"{record[0]} - {record[1]} - {record[2]}"
            ws1.write(i + 1, j + 1, record_str, fmt_normal)
        else:
            ws1.write(i + 1, j + 1, "", fmt_black)

# add overall record at bottom
    overall_row = 12
    ws1.write(overall_row, 0, "Overall", fmt_bold)

# calulate overall record
    for j, player in enumerate(sorted_players):
        total_wins = 0
        total_losses = 0
        total_ties = 0

        # Sum up all years for this player
        for year in yearly_records[player]:
            record = yearly_records[player][year]
            total_wins += record[0]
            total_losses += record[1]
            total_ties += record[2]

        ws1.write(overall_row, j + 1, f"{total_wins} - {total_losses} - {total_ties}", fmt_bold)



# onto ws2: yearly scores
logging.debug("Creating worksheet 2")

ws2 = workbook.add_worksheet("Regular Season Scores")
# get points for and against dict
player_scores_dict = fls.calculate_reg_averages(df1)[4]
player_career_avg = fls.calculate_reg_averages(df1)[2]
ws2.set_column_pixels(0, 0, 75, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws2.set_column_pixels(i, i, 150, fmt_bold)
for i in range(12):  # 11 years + header
    ws2.set_row_pixels(i, 24, fmt_bold)

# write headers
ws2.write(0, 0, "Year")
for i, name in enumerate(mapped_names):
    ws2.write(0, i + 1, name, fmt_bold)

# write year labels and records
for i, year in enumerate(range(2014, 2025)):
    ws2.write(i + 1, 0, year, fmt_bold)

    # write each player's record for that year
    for j, player in enumerate(sorted_players):
        if year in player_scores_dict[player]:
            total_for = player_scores_dict[player][year][0]
            total_against = player_scores_dict[player][year][1]
            scores_str = f"{total_for} - {total_against}"
            if scores_str == "0 - 0":
                ws2.write(i + 1, j + 1, "", fmt_black)
            else:
                ws2.write(i + 1, j + 1, scores_str, fmt_normal)
        else:
            ws2.write(i + 1, j + 1, "", fmt_black)

    ws2.write(12, 0, "Overall", fmt_bold)
    ws2.write(13, 0, "Average", fmt_bold)

# calulate overall record
    for j, player in enumerate(sorted_players):
        overall_points_for = 0
        overall_points_against = 0
        # Sum up all years for this player
        for year in player_scores_dict[player]:
            overall_points_for += player_scores_dict[player][year][0]
            overall_points_against += player_scores_dict[player][year][1]
    # write average and overall rows
        ws2.write(12, j + 1, f"{round(overall_points_for, 2)} - {round(overall_points_against, 2)}", fmt_bold)
        ws2.write(13, j + 1, f"{round(player_career_avg[player][0], 2)} - {round(player_career_avg[player][1], 2)}", fmt_bold)

# done with that



# ws3: h2h records
logging.debug("Creating worksheet 3")
ws3 = workbook.add_worksheet("Head to Head Record")
# get points for and against dict
h2h_records = fls.calculate_reg_records(df1)[0]

ws3.set_column_pixels(0, 0, 70, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws3.set_column_pixels(i, i, 70, fmt_bold)
for i in range(12):  # 11 years + header
    ws3.set_row_pixels(i, 70, fmt_bold)

# write headers
ws3.write(0, 0, "Year")
for i, name in enumerate(mapped_names):
    ws3.write(i + 1, 0, name, fmt_bold)
    ws3.write(0, i + 1, name, fmt_bold)

# write year labels and records
for i, p1 in enumerate(sorted_players):
    for j, p2 in enumerate(sorted_players):
        if p1 == p2:
            ws3.write(i + 1, j + 1, "---", fmt_black)  # Diagonal
        else:
            rec = h2h_records[p1][p2]
            ws3.write(i + 1, j + 1, f"{rec[0]} - {rec[1]} - {rec[2]}", fmt_normal)

# done



# ws4: h2h scores
logging.debug("Creating worksheet 4")
ws4 = workbook.add_worksheet("Head to Head Scores")
# get points for and against dict
h2h_scores = fls.calculate_reg_records(df1)[1]
h2h_count_reg = fls.calculate_reg_records(df1)[4]

ws4.set_column_pixels(0, 0, 100, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws4.set_column_pixels(i, i, 100, fmt_bold)
for i in range(12):  # 11 years + header
    ws4.set_row_pixels(i, 100, fmt_bold)

# write headers
ws4.write(0, 0, "Year", fmt_bold)
for i, name in enumerate(mapped_names):
    ws4.write(i + 1, 0, name, fmt_bold)
    ws4.write(0, i + 1, name, fmt_bold)

running_dict_player= {player:{opponent:[0,0] for opponent in sorted_players if opponent != player} for player in sorted_players}
# write year labels and records
for i, p1 in enumerate(sorted_players):
    for j, p2 in enumerate(sorted_players):
        if p1 == p2:
            ws4.write(i + 1, j + 1, "---", fmt_black)  # Diagonal
        else:
            avg_for = round(h2h_scores[p1][p2][0] / h2h_count_reg[p1][p2], 2)
            avg_against = round(h2h_scores[p1][p2][1] / h2h_count_reg[p1][p2], 2)
            ws4.write(i + 1, j + 1, f"{avg_for} - {avg_against}", fmt_normal)

# done


# ws5: 25 highest #
# decided to try doing this one with pandas because am learning it now and better than dictionaries
# (even though less comfortable for me)
logging.debug("Creating worksheet 5")

ws5 = workbook.add_worksheet('Highest Scores')
ws5.set_column_pixels(0, 0, 70, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws5.set_column_pixels(i, i, 150, fmt_bold)
for i in range(12):  # 11 years + header
    ws5.set_row_pixels(i, 24, fmt_bold)

highest_scores = fls.find_extreme_scores(df1)[0]

df_highest = pd.DataFrame(highest_scores)
df_highest['player'] = df_highest['player'].map(name_mapping)
df_highest['opponent'] = df_highest['opponent'].map(name_mapping)
df_highest = df_highest[['score', 'player', 'opponent', 'week', 'year']]
df_highest.columns = ['Score', 'Player', 'Opponent', 'Week', 'Year']

# make rank column
df_highest.insert(0, 'Rank', range(1, len(df_highest) + 1))
#  headers
headers = ['Rank', 'Score', 'Player', 'Opponent', 'Week', 'Year']
for col, header in enumerate(headers):
    ws5.write(0, col, header, fmt_bold)

# write data
for row_idx, score_dict in enumerate(highest_scores, start=1):
    ws5.write(row_idx, 0, row_idx, fmt_bold)  # Rank
    ws5.write(row_idx, 1, score_dict['score'], fmt_center)
    ws5.write(row_idx, 2, name_mapping[score_dict['player']], fmt_center)
    ws5.write(row_idx, 3, name_mapping.get(score_dict['opponent'], 'N/A'), fmt_center)
# I have literally no idea why I wrote this as a regex but its funny so why not?
    ws5_match = re.match(r'([^\d]+)(\d+)', score_dict['week'])
    ws5.write(row_idx, 4, f"{ws5_match.group(1)} {ws5_match.group(2)}", fmt_normal)# kind of tired of coding and having a brain fart couldn't remember a better way to do it
    ws5.write(row_idx, 5, score_dict['year'], fmt_center)


ws5.set_column_pixels(7, 10, 150, fmt_center)
rankings_ws5_headers = ["Player", "Appearances", "Opponent Appearances", "Rank Points"]
for i, header in enumerate(rankings_ws5_headers):
    ws5.write(0, i + 7, header, fmt_bold)

# count appearances for each player across all highest scores
player_appearance_count = {player: 0 for player in sorted_players}
opponent_appearance_count = {player: 0 for player in sorted_players}

for score_dict in highest_scores:
    player_appearance_count[score_dict['player']] += 1
    opponent_appearance_count[score_dict['opponent']] += 1

# write the rankings data for each player
for i, p1 in enumerate(sorted_players, start=1):
    ws5.write(i, 7, name_mapping[p1], fmt_bold)
    ws5.write(i, 8, player_appearance_count[p1], fmt_center)
    ws5.write(i, 9, opponent_appearance_count[p1], fmt_center)

    # calculate rank points (25 points for 1st place, 24 for 2nd, etc.), subtract if opponent of high_score
    rank_points = 0
    for rank, score_dict in enumerate(highest_scores, start=1):
        if score_dict['player'] == p1:
            rank_points += (26 - rank)

    ws5.write(i, 10, rank_points, fmt_center)

# done

### ws6: 25 lowest
logging.debug("Creating worksheet 6")

ws6 = workbook.add_worksheet('Lowest Scores')

lowest_scores = fls.find_extreme_scores(df1)[1]
lowest_scores = lowest_scores[::-1]

df_lowest = pd.DataFrame(lowest_scores)
df_lowest['player'] = df_lowest['player'].map(name_mapping)
df_lowest['opponent'] = df_lowest['opponent'].map(name_mapping)
df_lowest = df_lowest[['score', 'player', 'opponent', 'week', 'year']]
df_lowest.columns = ['Score', 'Player', 'Opponent', 'Week', 'Year']

# make rank column
df_lowest.insert(0, 'Rank', range(1, len(df_lowest) + 1))
#  headers
headers = ['Rank', 'Score', 'Player', 'Opponent', 'Week', 'Year']
for col, header in enumerate(headers):
    ws6.write(0, col, header, fmt_bold)

# write data
for row_idx, score_dict in enumerate(lowest_scores, start=1):
    ws6.write(row_idx, 0, row_idx, fmt_bold)  # Rank
    ws6.write(row_idx, 1, score_dict['score'], fmt_center)
    ws6.write(row_idx, 2, name_mapping[score_dict['player']], fmt_center)
    ws6.write(row_idx, 3, name_mapping.get(score_dict['opponent'], 'N/A'), fmt_center)
    # I have literally no idea why I wrote this as a regex but its funny so why not?
    ws6_match = re.match(r'([^\d]+)(\d+)', score_dict['week'])
    ws6.write(row_idx, 4, f"{ws6_match.group(1)} {ws6_match.group(2)}", fmt_normal)
    # kind of tired of coding and having a brain fart couldn't remember a better way to do it
    ws6.write(row_idx, 5, score_dict['year'], fmt_center)

ws6.set_column_pixels(7, 10, 150, fmt_center)
rankings_ws6_headers = ["Player", "Appearances", "Opponent Appearances", "Rank Points"]
for i, header in enumerate(rankings_ws6_headers):
    ws6.write(0, i + 7, header, fmt_bold)

# count appearances for each player across all lowest scores
player_appearance_count = {player: 0 for player in sorted_players}
opponent_appearance_count = {player: 0 for player in sorted_players}

for score_dict in lowest_scores:
    player_appearance_count[score_dict['player']] += 1
    opponent_appearance_count[score_dict['opponent']] += 1

# write the rankings data for each player
for i, p1 in enumerate(sorted_players, start=1):
    ws6.write(i, 7, name_mapping[p1], fmt_bold)
    ws6.write(i, 8, player_appearance_count[p1], fmt_center)
    ws6.write(i, 9, opponent_appearance_count[p1], fmt_center)

    # calculate rank points
    # want the fewest (25 points for 1st place, 24 for 2nd, etc.), lose 25 for being opponent for first, etc
    rank_points = 0
    for rank, score_dict in enumerate(lowest_scores, start=1):
        if score_dict['player'] == p1:
            rank_points += (26 - rank)  # 25 for 1st, 24 for 2nd, etc.


    ws6.write(i, 10, rank_points, fmt_center)
# done

### ws7: 25 largest victories
logging.debug("Creating worksheet 7")
ws7 = workbook.add_worksheet("Largest Victories")

largest_wins = fls.find_extreme_scores(df1)[4]

df_largest_wins = pd.DataFrame(largest_wins)
df_largest_wins['player'] = df_largest_wins['player'].map(name_mapping)
df_largest_wins['opponent'] = df_largest_wins['opponent'].map(name_mapping)
df_largest_wins = df_largest_wins[['margin', 'player', 'opponent', 'week', 'year']]
df_largest_wins.columns = ['Margin', 'Player', 'Opponent', 'Week', 'Year']

# make rank column
df_largest_wins.insert(0, 'Rank', range(1, len(df_largest_wins) + 1))
#  headers
headers = ['Rank', 'Margin', 'Player', 'Opponent', 'Week', 'Year']
for col, header in enumerate(headers):
    ws7.write(0, col, header, fmt_bold)

# write data
for row_idx, score_dict in enumerate(largest_wins, start=1):
    ws7.write(row_idx, 0, row_idx, fmt_bold)  # Rank
    ws7.write(row_idx, 1, score_dict["margin"], fmt_center)
    ws7.write(row_idx, 2, name_mapping[score_dict['player']], fmt_center)
    ws7.write(row_idx, 3, name_mapping.get(score_dict['opponent'], 'N/A'), fmt_center)
    # I have literally no idea why I wrote this as a regex but its funny so why not?
    ws7_match = re.match(r'([^\d]+)(\d+)', score_dict['week'])
    ws7.write(row_idx, 4, f"{ws7_match.group(1)} {ws7_match.group(2)}", fmt_normal)
    # kind of tired of coding and having a brain fart couldn't remember a better way to do it
    ws7.write(row_idx, 5, score_dict['year'], fmt_center)

ws7.set_column_pixels(7, 10, 150, fmt_center)
rankings_ws7_headers = ["Player", "Appearances", "Opponent Appearances", "Rank Points"]
for i, header in enumerate(rankings_ws7_headers):
    ws7.write(0, i + 7, header, fmt_bold)

# count appearances for each player across all largest margins of vitory
player_appearance_count = {player: 0 for player in sorted_players}
opponent_appearance_count = {player: 0 for player in sorted_players}

for score_dict in largest_wins:
    player_appearance_count[score_dict['player']] += 1
    opponent_appearance_count[score_dict['opponent']] += 1

# write the rankings data for each player
for i, player in enumerate(sorted_players, start=1):
    ws7.write(i, 7, name_mapping[p1], fmt_bold)
    ws7.write(i, 8, player_appearance_count[p1], fmt_center)
    ws7.write(i, 9, opponent_appearance_count[p1], fmt_center)

    # calculate rank points
    # want the fewest (25 points for 1st place, 24 for 2nd, etc.), lose 25 for being opponent for first, etc
    rank_points = 0
    for rank, score_dict in enumerate(largest_wins, start=1):
        if score_dict["player"] == player:
            rank_points += (26 - rank)  # 25 for 1st, 24 for 2nd, etc.
        if score_dict["opponent"] == player:
            rank_points -= (26 - rank)
    ws7.write(i, 10, rank_points, fmt_center)
# done


# ws8: made playoffs #
logging.debug("Creating worksheet 8")

ws8 = workbook.add_worksheet("Made Playoffs")
made_playoffs_dict = fls.calculate_playoff_makes_misses(df1)[0]
year_postseason_outlook = fls.calculate_playoff_makes_misses(df1)[3]
made_playoffs_list = {player:[] for player in players}

ws8.set_column_pixels(0, 0, 90, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws8.set_column_pixels(i, i, 90, fmt_center)
for i in range(12):  # 11 years + header
    ws8.set_row_pixels(i, 24, fmt_center)

# write headers
ws8.write(0, 0, "Year", fmt_bold)
for i, name in enumerate(mapped_names):
    ws8.write(0, i + 1, name, fmt_bold)

# write year labels and records
for i, year in enumerate(range(2014, 2025)):
    ws8.write(i + 1, 0, year, fmt_bold)
    for j, player in enumerate(sorted_players):
        if player in year_postseason_outlook[year][0]:
            ws8.write(i + 1, j + 1, str(made_playoffs_dict[player][year]), fmt_bold)
            made_playoffs_list[player].append(made_playoffs_dict[player][year])
        elif player in year_postseason_outlook[year][1] or player in year_postseason_outlook[year][2]:
            ws8.write(i + 1, j + 1, str(made_playoffs_dict[player][year]), fmt_normal)
        else:
            ws8.write(i + 1, j + 1, "", fmt_black)
    ws8.write(i + 1, 11, len(year_postseason_outlook[year][0]), fmt_bold)

ws8.write(12,0,"Total", fmt_bold)
for j, player in enumerate(sorted_players):
    num_seasons = 11
    if player == "Rockmael" or player == "Jackson":
        num_seasons = 9
    ws8.write(12, j+1, f"{Counter(made_playoffs_list[player])[True]} / {num_seasons}")
ws8.write(0, 11, "Check", fmt_bold)

# done


# ws9: made Loser's bowl #
logging.debug("Creating worksheet 9")

ws9 = workbook.add_worksheet("Made Losers Bowl")
made_losers_dict = fls.calculate_playoff_makes_misses(df1)[2]
made_losers_list = {player: [] for player in sorted_players}

ws9.set_column_pixels(0, 0, 90, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws9.set_column_pixels(i, i, 90, fmt_center)
for i in range(12):  # 11 years + header
    ws9.set_row_pixels(i, 24, fmt_center)

# write headers
ws9.write(0, 0, "Year", fmt_bold)
for i, name in enumerate(mapped_names):
    ws9.write(0, i + 1, name, fmt_bold)

# write year labels and records
for i, year in enumerate(range(2014, 2025)):
    ws9.write(i + 1, 0, year, fmt_bold)
    for j, player in enumerate(sorted_players):
        if player in year_postseason_outlook[year][2]:
            ws9.write(i + 1, j + 1, str(made_losers_dict[player][year]), fmt_bold)
            made_losers_list[player].append(made_losers_dict[player][year])
        elif (player in year_postseason_outlook[year][1]) or (player in year_postseason_outlook[year][0]):
            ws9.write(i + 1, j + 1, str(made_losers_dict[player][year]), fmt_normal)
        else:
            ws9.write(i + 1, j + 1, "", fmt_black)
    ws9.write(i + 1, 11, len(year_postseason_outlook[year][2]), fmt_bold)

ws9.write(12,0,"Total", fmt_bold)
for j, player in enumerate(sorted_players):
    num_seasons = 11
    if player == "Rockmael" or player == "Jackson":
        num_seasons = 9
    ws9.write(12, j+1, f"{Counter(made_losers_list[player])[True]} / {num_seasons}")
ws9.write(0, 11, "Check", fmt_bold)


# ws10 avg regular season finish #
logging.debug("Creating worksheet 10")

ws10 = workbook.add_worksheet("Average Regular Season Finish")
standings_dictionary = fls.create_standings(df1)
list_of_reg_finishes = {player: [] for player in players}

ws10.set_column_pixels(0, 0, 90, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws10.set_column_pixels(i, i, 90, fmt_center)
for i in range(12):  # 11 years + header
    ws10.set_row_pixels(i, 24, fmt_center)

# write headers
ws10.write(0, 0, "Year", fmt_bold)
for i, name in enumerate(mapped_names):
    ws10.write(0, i + 1, name, fmt_bold)

ws10.write(12, 0, "Average", fmt_bold)
ws10.write(13, 0, "Most Common", fmt_bold)

# write year labels and records
for i, year in enumerate(range(2014, 2025)):
    ws10.write(i + 1, 0, year, fmt_bold)
    for j, player in enumerate(sorted_players):
        if player in standings_dictionary[year]:
            ws10.write(i + 1, j + 1, str(standings_dictionary[year][player]["RegularSeasonRank"]), fmt_normal)
            list_of_reg_finishes[player].append(standings_dictionary[year][player]["RegularSeasonRank"])
        else:
            ws10.write(i + 1, j + 1, "", fmt_black)

for j, player in enumerate(sorted_players):
    ws10.write(12, j + 1, round(sum(list_of_reg_finishes[player])/ len(list_of_reg_finishes[player]), 2),fmt_bold)
    ws10.write(13, j + 1, Counter(list_of_reg_finishes[player]).most_common(1)[0][0], fmt_bold)

# done


# ws11 avg final standings finish #
logging.debug("Creating worksheet 11")

ws11 = workbook.add_worksheet("Average Final Finish")
standings_dictionary = fls.create_standings(df1)
list_of_final_finishes = {player: [] for player in players}

ws11.set_column_pixels(0, 0, 90, fmt_center)
for i in range(1, len(sorted_players) + 1):
    ws11.set_column_pixels(i, i, 90, fmt_center)
for i in range(12):  # 11 years + header
    ws11.set_row_pixels(i, 24, fmt_center)

# write headers
ws11.write(0, 0, "Year", fmt_bold)
for i, name in enumerate(mapped_names):
    ws11.write(0, i + 1, name, fmt_bold)

ws11.write(12, 0, "Average", fmt_bold)
ws11.write(13, 0, "Most Common", fmt_bold)

# write year labels and records
for i, year in enumerate(range(2014, 2025)):
    ws11.write(i + 1, 0, year, fmt_bold)
    for j, player in enumerate(sorted_players):
        if player in standings_dictionary[year]:
            ws11.write(i + 1, j + 1, str(standings_dictionary[year][player]["PlayoffRank"]), fmt_normal)
            list_of_final_finishes[player].append(standings_dictionary[year][player]["PlayoffRank"])
        else:
            ws11.write(i + 1, j + 1, "", fmt_black)

for j, player in enumerate(sorted_players):
    ws11.write(12, j + 1, round(sum(list_of_final_finishes[player])/ len(list_of_final_finishes[player]), 2),fmt_bold)
    ws11.write(13, j + 1, Counter(list_of_final_finishes[player]).most_common(1)[0][0], fmt_bold)


# done!

workbook.close()

