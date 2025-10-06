import csv
from pickletools import stringnl_noescape
from re import match
import os
import pandas as pd
import xlsxwriter
from io import StringIO

from numpy.f2py.crackfortran import word_pattern

from constants import standings_directory

# basic useful information
players = ["Aaron", "Adam", "Jackson", "Marder", "Oliver", "Rockmael", "Saxe", "Steven", "Todd", "Whyte"]
players_pre_2016 = ["Aaron", "Adam", "Marder", "Oliver", "Saxe", "Steven", "Todd", "Whyte"]
name_mapping = {
    "Marder": "Liam M", "Whyte": "Liam W", "Aaron": "Aaron", "Adam": "Adam",
    "Todd": "Todd", "Rockmael": "Kevin R", "Jackson": "Jackson",
    "Saxe": "Kevin S", "Oliver": "Oliver", "Steven": "Steven"
}

# league structure by year
league_set_up = {}
for year in range(2014,2025):
    number_of_teams = 10
    players_list = players
    if year < 2016:
        number_of_teams = 8
        players_list = players_pre_2016
    reg_season_weeks = 15
    if year < 2021:
        reg_season_weeks = 14
    league_set_up[year] = {"teams": number_of_teams, "reg_season_weeks":reg_season_weeks, "players":players_list}

# more helpful stuff:
number_of_seasons = {}
for player in players:
    seasons = 10
    if player not in players_pre_2016:
        seasons = 8
    number_of_seasons[player] = seasons

# dataframe representing all stats for the league
fantasy_df = pd.read_excel("fantasy_league_stats.xlsx")

def get_year_rows(year, df):
# getting the rows in the dataframe that represent each year/season
    years = df["Year"].dropna() # exclude empty cells
    year_index = years[years == year].index.tolist()[0] # index of year row (is at the top row for each season)
    num_matchups = league_set_up[year]["teams"] // 2  # number of matchups per week = no. players/2
    return list(range(year_index, year_index + num_matchups)) # returns list of indices for the rows that make up the season


def extract_reg_matchup_data(df):
    player_scores_dict = {player: {} for player in players}
    player_opponents_dict = {player: {} for player in players}

    for year, set_up in league_set_up.items():
        year_rows = get_year_rows(year, df)
        if not year_rows:
            continue

        teams = set_up["teams"]
        active_players = set_up["players"]
        reg_weeks = set_up["reg_season_weeks"]

        for player in active_players:
            player_scores_dict[player][year] = {}
            player_opponents_dict[player][year] = {}

            for week in range(1, reg_weeks + 1):
                week_col = f"Week{week}"
                score_col = f"Score{week}"
                week_col_2 = f"Week{week}.5"
                score_col_2 = f"Score{week}.5"

            # combine both matchups for the week
                matchups = list(df.loc[year_rows, week_col].values) + list(df.loc[year_rows, week_col_2].values)
                scores = list(df.loc[year_rows, score_col].values) + list(df.loc[year_rows, score_col_2].values)

            # check that both scores and matchup are not null
                valid_data = [(matchup, score) for matchup, score in zip(matchups, scores) if pd.notna(matchup) and pd.notna(score)]
                matchups = [matchup for matchup, score in valid_data]
                scores = [score for matchup, score in valid_data]

                if player in matchups:
                    idx = matchups.index(player)
                    player_scores_dict[player][year][f"Week{week}"] = scores[idx]

                # get opponent
                    num_matchups = teams // 2
                    opp_idx = (idx + num_matchups) % teams
                    player_opponents_dict[player][year][f"Week{week}"] = matchups[opp_idx]
    return player_scores_dict, player_opponents_dict

def extract_postseason_matchup_data(df):
    player_scores_dict = {player: {} for player in players}
    player_opponents_dict = {player: {} for player in players}

    for year, set_up in league_set_up.items():
        year_rows = get_year_rows(year, df)
        if not year_rows:
            continue

        teams = set_up["teams"]
        active_players = set_up["players"]
        playoff_weeks = set_up["reg_season_weeks"]

        for player in active_players:
            player_scores_dict[player][year] = {}
            player_opponents_dict[player][year] = {}

            for week in range(playoff_weeks+1, playoff_weeks + 3):
                week_col = f"Week{week}"
                score_col = f"Score{week}"
                week_col_2 = f"Week{week}.5"
                score_col_2 = f"Score{week}.5"

            # combine both matchups for the week
                matchups = list(df.loc[year_rows, week_col].values) + list(df.loc[year_rows, week_col_2].values)
                scores = list(df.loc[year_rows, score_col].values) + list(df.loc[year_rows, score_col_2].values)

                if player in matchups:
                    idx = matchups.index(player)
                    player_scores_dict[player][year][f"Week{week}"] = scores[idx]

                # get opponent
                    num_matchups = teams // 2
                    opp_idx = (idx + num_matchups) % teams
                    player_opponents_dict[player][year][f"Week{week}"] = matchups[opp_idx]
                else:
                    player_scores_dict[player][year][f"Week{week}"] = "NA"
                    player_opponents_dict[player][year][f"Week{week}"] = "NA"
    return player_scores_dict, player_opponents_dict



def calculate_reg_records(df):
    # returns a list of dictionaries: [h2h_records, h2h_scores, yearly_records, weekly_records]
    # h2h_records format {player:{opponent:[W,L,T]}}
    # h2h_scores format {player:{opponent:[points_for, points_against]}}
    # yearly_records format: {player:{year:[W,L,T]}}
    # weekly_records format: {player:{year:{week:[W,L,T]}}} where each [W,L,T] represents player's record at end of specified week

    extracted_matchups = extract_reg_matchup_data(df)
    player_scores_dict = extracted_matchups[0]
    player_opponents_dict = extracted_matchups[1]
    h2h_records = {player: {} for player in players}
    h2h_scores = {player: {} for player in players}
    h2h_count_reg = {player: {} for player in players}

    for player in players:
        for opponent in players:
            if opponent != player:
                h2h_records[player][opponent] = [0, 0, 0]  # [W,L,T]
                h2h_scores[player][opponent] = [0, 0]  # pf-pa
                h2h_count_reg[player][opponent] = 0

    yearly_records = {player: {} for player in players}
    weekly_records = {player: {year:{} for year in league_set_up} for player in players}

    for player in player_scores_dict:
        for year in player_scores_dict[player]:
            yearly_records[player][year] = [0, 0, 0]  # [W, L, T]
            for week in player_scores_dict[player][year]:
                if week not in player_opponents_dict[player][year]:
                    continue
                opponent = player_opponents_dict[player][year][week]
                player_score = player_scores_dict[player][year][week]
                opp_score = player_scores_dict[opponent][year][week]
            # load scores into h2h scores dict
                h2h_count_reg[player][opponent] += 1
                h2h_scores[player][opponent][0] += player_score
                h2h_scores[player][opponent][1] += opp_score

            # determine who won and update h2h records dict and records dicts accordingly
                if player_score > opp_score:
                    h2h_records[player][opponent][0] += 1
                    yearly_records[player][year][0] += 1

                elif player_score < opp_score:
                    h2h_records[player][opponent][1] += 1
                    yearly_records[player][year][1] += 1
                else:
                    h2h_records[player][opponent][2] += 1
                    yearly_records[player][year][2] += 1
                weekly_records[player][year][week] = [yearly_records[player][year][0], yearly_records[player][year][1], yearly_records[player][year][2]]
    return h2h_records, h2h_scores, yearly_records, weekly_records, h2h_count_reg

def create_standings(df):
    # returns a dictionary standings_dictionary
    # return format = {year:{player:{

# calling function here so don't have to do it over and over in loops
    yearly_records_dict = calculate_reg_records(df)[2]
    reg_scores_dict = extract_reg_matchup_data(df)[0]

#
    standings_dictionary = {year: {player:{} for player in league_set_up[year]["players"]} for year in league_set_up}
    
# both Kevins are excluded from the following dictionary because they both have the same "ManagerName"
# since they both have the same name on NFL fantasy, have to find a way to double triple check that it's the right Kevin
    player_manager_translation_dict = {"Marder": "Liam", "Whyte": "Liam WHYTE", "Aaron": "Aaron", "Adam": "Adam","Todd": "Todd", "Jackson": "Jackson", "Oliver": "Oliver", "Steven": "Steven"}
    manager_player_translation_dict = {value: key for key, value in player_manager_translation_dict.items()}

    for year in range(2014,2025): # start iterating through years

    # defining the kevins' records and scores for the year so can check the csv against them
    # and so I don't have to call the function over and over again (way too slow)
        saxe_record = yearly_records_dict["Saxe"][year]
        saxe_record_combined = f"{saxe_record[0]}-{saxe_record[1]}-{saxe_record[2]}"
        saxe_score = sum([reg_scores_dict["Saxe"][year][f"Week{week}"] for week in
                              range(1, league_set_up[year]["reg_season_weeks"] + 1)])

    # create pd df for csv for that year's standings
        standings_csv = pd.read_csv(os.path.join(
                r"C:\Users\aaron\PycharmProjects\oct25_fantasy_stats_with_scraping\output\2457715-history-standings",
                f"{year}.csv"))
        for _, row in standings_csv.iterrows(): # don't actually need index
            manager = row["ManagerName"]
            # straightforward to create dictionary for non-kevin players
            if manager in manager_player_translation_dict:
                player = manager_player_translation_dict[manager]
                for field in standings_csv.columns.tolist():
                    standings_dictionary[year][player][field] = row[field]

        #!!!# the kevin contingency!!!#
            elif manager == "Kevin" and year < 2016: # before 2016 it could only be K Saxe
                for field in standings_csv.columns.tolist():
                    standings_dictionary[year]["Saxe"][field] = row[field]

            else:
                # just need a place to define these vars so they're accessible below
                rockmael_record = yearly_records_dict["Rockmael"][year]
                rockmael_score = sum([reg_scores_dict["Rockmael"][year][f"Week{week}"] for week in
                                      range(1, league_set_up[year]["reg_season_weeks"] + 1)])
                rockmael_record_combined = f"{rockmael_record[0]}-{rockmael_record[1]}-{rockmael_record[2]}"


            # checking to see if both Kevin's records are equal, and row["PointsFor"] does not match either of their points (SHOULDNT HAPPEN!)
                if (manager == "Kevin") and (rockmael_record == saxe_record) \
                        and not (float(row["PointsFor"].replace(",", "")) == round(saxe_score, 2) or float(row["PointsFor"].replace(",", "")) == round(rockmael_score,2)):
                    raise Exception("kill me now")
            # matches K Saxe record?
                elif row["Record"] == saxe_record_combined:
                    for field in standings_csv.columns.tolist():
                        standings_dictionary[year]["Saxe"][field] = row[field]
            # matches K Rock record?
                elif row["Record"] == rockmael_record_combined:
                    for field in standings_csv.columns.tolist():
                        standings_dictionary[year]["Rockmael"][field] = row[field]
    # double-checking that each kevin appears once for
        if year >= 2016 and ("Rockmael" not in standings_dictionary[year]):
            raise Exception("this is even worse")
        elif "Saxe" not in standings_dictionary[year]:
            raise Exception("Missing KSaxe")
    return standings_dictionary

def calculate_playoff_makes_misses(df):
    made_playoffs_dict = {player: {} for player in players}
    made_consolation_bracket_dict = {player: {} for player in players}
    made_losers_bowl_dict = {player: {} for player in players}
    standings = create_standings(df)
    for player in players:
        for year in range(2014,2025):

        # have to check whether player is in standings for year to avoid error from rockmaels
            if player in standings[year]:
                made_playoffs_dict[player][year] = False
                made_consolation_bracket_dict[player][year] = False
                made_losers_bowl_dict[player][year] = False
                if standings[year][player]["RegularSeasonRank"] >= 4:
                    made_playoffs_dict[player][year] = True
                elif standings[year][player]["RegularSeasonRank"] >= 8:
                    made_consolation_bracket_dict[player][year] = True
                else:
                    made_losers_bowl_dict[player][year] = True
    return made_playoffs_dict, made_consolation_bracket_dict, made_losers_bowl_dict

def calculate_playoff_records(df):

    record_in_playoffs = {player: [0,0,0] for player in players}
    h2h_record_playoffs = {player:{} for player in players}
    h2h_scores_playoffs = {player:{} for player in players}
    h2h_count_playoffs = {player:{} for player in players}

    for player in players:
        for opponent in players:
            if opponent != player:
                h2h_record_playoffs[player][opponent] = [0,0,0]
                h2h_scores_playoffs[player][opponent] = [0,0]
                h2h_count_playoffs[player][opponent] = 0

    postseason_matchups = extract_postseason_matchup_data(df)
    player_scores_dict = postseason_matchups[0]
    player_opponents_dict = postseason_matchups[1]

    for player in player_scores_dict:
        for year in player_scores_dict[player]:
            if calculate_playoff_makes_misses(df)[0][player][year]:
                for week in player_scores_dict[player][year]:
                    opponent = player_opponents_dict[player][year][week]
                    player_score = player_scores_dict[player][year][week]
                    opp_score = player_scores_dict[opponent][year][week]
                # load scores into h2h scores dict
                    h2h_scores_playoffs[player][opponent][0] += player_score
                    h2h_scores_playoffs[player][opponent][1] += opp_score
                    h2h_count_playoffs[player][opponent] += 1
    
                # determine who won and update h2h records dict and records dicts accordingly
                    if player_score > opp_score:
                        h2h_record_playoffs[player][opponent][0] += 1
                        record_in_playoffs[player][0] += 1
    
                    elif player_score < opp_score:
                        h2h_record_playoffs[player][opponent][1] += 1
                        record_in_playoffs[player][1] += 1
                    else:
                        h2h_record_playoffs[player][opponent][1] += 1
                        record_in_playoffs[player][1] += 1
    return record_in_playoffs, h2h_record_playoffs, h2h_scores_playoffs, h2h_count_playoffs


def calculate_postseason_records(df):
    record_in_postseason = {player: [0, 0, 0] for player in players}
    h2h_record_postseason = {player: {} for player in players}
    h2h_scores_postseason = {player: {} for player in players}
    h2h_count_postseason = {player: {} for player in players}

    for player in players:
        for opponent in players:
            if opponent != player:
                h2h_record_postseason[player][opponent] = [0, 0, 0]
                h2h_scores_postseason[player][opponent] = [0, 0]
                h2h_count_postseason[player][opponent] = 0

    postseason_matchups = extract_postseason_matchup_data(df)
    player_scores_dict = postseason_matchups[0]
    player_opponents_dict = postseason_matchups[1]

    for player in player_scores_dict:
        for year in player_scores_dict[player]:
             for week in player_scores_dict[player][year]:
                    opponent = player_opponents_dict[player][year][week]
                    player_score = player_scores_dict[player][year][week]
                    opp_score = player_scores_dict[opponent][year][week]
                    # load scores into h2h scores dict
                    h2h_scores_postseason[player][opponent][0] += player_score
                    h2h_scores_postseason[player][opponent][1] += opp_score
                    h2h_count_postseason[player][opponent] += 1

                    # determine who won and update h2h records dict and records dicts accordingly
                    if player_score > opp_score:
                        h2h_record_postseason[player][opponent][0] += 1
                        record_in_postseason[player][0] += 1

                    elif player_score < opp_score:
                        h2h_record_postseason[player][opponent][1] += 1
                        record_in_postseason[player][1] += 1
                    else:
                        h2h_record_postseason[player][opponent][1] += 1
                        record_in_postseason[player][1] += 1
    return record_in_postseason, h2h_record_postseason, h2h_scores_postseason, h2h_count_postseason
        


def calculate_reg_averages(df):
    # returns avg_score_h2h, avg_score_season, avg_score_career, total_games
    # [0] = avg_score_h2h format: avg_score_h2h[player][opponent] = [points_for, points_against]
    # [1] = avg_score_season: avg_score_h2h[player][year] = [points_for, points_against]

    # calculating point records
    reg_records = calculate_reg_records(df)
    extracted_reg = extract_reg_matchup_data(df)
    h2h_records = reg_records[0]
    h2h_scores = reg_records[1]
    avg_score_season = {player: {year: [0,0] for year in league_set_up} for player in players}
    avg_score_career = {player: [0,0] for player in players}
    total_games = {player: 0 for player in players}
    avg_score_h2h = {player: {} for player in players}
    total_score = {player: {year: [0,0] for year in league_set_up} for player in players}

    for player in players:
        total_points_for = round(sum(h2h_scores[player][opp][0] for opp in h2h_scores[player]),2)
        total_points_against = round(sum(h2h_scores[player][opp][1] for opp in h2h_scores[player]), 2)
        games = sum(sum(h2h_records[player][opp]) for opp in h2h_records[player])
        total_games[player] = games
        avg_score_career[player] = [total_points_for/float(games), total_points_against / games]

    # now adding avg h2h scores
        for opponent in h2h_scores[player]:
            avg_score_h2h[player][opponent] = [0, 0]
            if opponent != player:
                avg_score_h2h[player][opponent][0] = round(h2h_scores[player][opponent][0]/reg_records[4][player][opponent], 2)
                avg_score_h2h[player][opponent][1] = round(h2h_scores[opponent][player][0]/reg_records[4][player][opponent],2)
    # just in case there's an error I safe guarded here


    # do average score for season and h2h
        for year in extracted_reg[0][player]:
            running_count_for = 0
            running_count_against = 0
            for week in extracted_reg[0][player][year]:
                running_count_for += extracted_reg[0][player][year][week]
                running_count_against += extracted_reg[0][extracted_reg[1][player][year][week]][year][week]
            total_score[player][year][0] = round(running_count_for, 2)
            total_score[player][year][1] = round(running_count_against,2)
            avg_score_season[player][year][0] = round(running_count_for/league_set_up[year]["reg_season_weeks"], 2)
            avg_score_season[player][year][1] = round(running_count_against/league_set_up[year]["reg_season_weeks"], 2)

    return avg_score_h2h, avg_score_season, avg_score_career, total_games, total_score

def find_extreme_scores(df):
    # twenty_five_highest_scores_overall, twenty_five_lowest_scores_overall in form: [{"score": score,"player": player,"opponent": opponent, "week": week,"year": year},...]
    # ten_highest_playoff_scores and ten_lowest_playoff_scores in form [{"score": score,"player": player,"opponent": opponent, "week": week,"year": year},...]
    # player_outlier_scores in form: {player: [{"score": score,"player": player,"opponent": opponent, "week": week,"year": year},...]}
    # 0 = 25_highest, 1 = 25_lowest, 2 = 10_highest, 3 = 10_lowest, 4 = player_outlier_scores

# add all scores to a list
    all_scores = []
    all_scores_per_player = {player:[] for player in players}
    all_playoff_scores = []
# get external sources
    reg_season_scores =  extract_reg_matchup_data(df)[0]
    post_season_scores = extract_postseason_matchup_data(df)[0]
    reg_season_opponents = extract_reg_matchup_data(df)[1]
    post_season_opponents = extract_postseason_matchup_data(df)[1]
    made_playoffs_dict = calculate_playoff_makes_misses(df)[0]


    for player in reg_season_scores:
        for year in reg_season_scores[player]:
            for week, score in reg_season_scores[player][year].items():
                opponent = reg_season_opponents[player][year].get(week)
                all_scores.append({"score": score,"player": player,"opponent": opponent,"week": week,"year": year})
                all_scores_per_player[player].append({"score": score,"player": player,"opponent": opponent,"week": week,"year": year})
    for player in post_season_scores:
        for year in post_season_scores[player]:
            for week, score in post_season_scores[player][year].items():
                if score == "NA":
                    continue
                opponent = post_season_opponents[player][year].get(week)
                all_scores.append({"score": score, "player": player, "opponent": opponent, "week": week, "year": year})
                all_scores_per_player[player].append(
                    {"score": score, "player": player, "opponent": opponent, "week": week, "year": year})
            # check if player made playoffs to see if this is a playoff score
                if made_playoffs_dict[player][year]:
                    all_playoff_scores.append({"score": score, "player": player, "opponent": opponent, "week": week, "year": year})



    sorted_scores = sorted(all_scores, key=lambda x: x["score"], reverse=True)
    twenty_five_highest_scores_overall = sorted_scores[:25]
    twenty_five_lowest_scores_overall = sorted_scores[-25:]
    ten_highest_playoff_scores = sorted(all_playoff_scores, key=lambda x: x["score"],reverse=True)[:10]
    ten_lowest_playoff_scores = sorted(all_playoff_scores, key=lambda x: x["score"], reverse=True)[-10:]
    player_outlier_scores = {player: [sorted(all_scores_per_player[player], key=lambda x: x["score"],reverse=True)[:10], sorted(all_scores_per_player[player], key=lambda x: x["score"],reverse=True)[-10:]] for player in players}


    return twenty_five_highest_scores_overall, twenty_five_lowest_scores_overall, ten_highest_playoff_scores, ten_lowest_playoff_scores, player_outlier_scores

# !! # COMPLETE UNTIL HERE # !! #

