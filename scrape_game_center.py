import csv
import os
from bs4 import BeautifulSoup as bs
from urllib.request import urlopen
import re
import requests
from cookie_string import cookies
from utils import get_number_of_owners, setup_output_folders
from constants import leagueID, leagueStartYear, leagueEndYear, gamecenter_directory


# teams that don't fill all their starting roster spots for a week will have a longer bench
# the more roster spots left unfilled, the more bench players that team will have
# this method gets the teamid of the team with the longest bench for the week as well as the length of their bench
def get_longest_bench(week):
    longest_bench_data = [0, 0]
    for i in range(1, number_of_owners + 1):
        # For 2015, add the gameCenterTab and trackType parameters to load Simple Box Score
        if season == '2015':
            url = f'https://fantasy.nfl.com/league/{leagueID}/history/{season}/teamgamecenter?gameCenterTab=track&teamId={i}&trackType=sbs&week={week}'
        else:
            url = f'https://fantasy.nfl.com/league/{leagueID}/history/{season}/teamgamecenter?teamId={i}&week={week}'

        page = requests.get(url, cookies=cookies)
        soup = bs(page.text, 'html.parser')
        print(i)

        # Try to find the bench table - handle cases where it doesn't exist
        bench_table = soup.find('div', id='tableWrapBN-1')

        if bench_table is None:
            # If the bench table doesn't exist, try alternative selectors
            print(f"  Warning: Bench table not found for team {i}, trying alternatives...")

            # Try finding by class pattern
            bench_table = soup.find('div', id=re.compile('tableWrapBN'))

            if bench_table is None:
                # If still not found, try to find any bench section
                bench_table = soup.find('div', class_=re.compile('bench'))

        if bench_table is not None:
            bench_length = len(bench_table.find_all('td', class_='playerNameAndInfo'))
        else:
            # If we still can't find it, use your league's actual bench size
            print(f"  Warning: Could not find bench for team {i}, using default size of 6")
            bench_length = 6

        if (bench_length > longest_bench_data[0]):
            longest_bench_data = [bench_length, i]

    return longest_bench_data


# generates the header for the csv file for the week
# different weeks can have different headers if players do not fill all their starting roster spots
def get_header(week, longest_bench_teamID):
    # For 2015, add the gameCenterTab and trackType parameters to load Simple Box Score
    if season == '2015':
        url = f"https://fantasy.nfl.com/league/{leagueID}/history/{season}/teamgamecenter?gameCenterTab=track&teamId={longest_bench_teamID}&trackType=sbs&week={week}"
    else:
        url = f"https://fantasy.nfl.com/league/{leagueID}/history/{season}/teamgamecenter?teamId={longest_bench_teamID}&week={week}"

    page = requests.get(url, cookies=cookies)
    html = page.text
    page.close()
    soup = bs(html, 'html.parser')  # uses the page of the teamID with the longest bench to generate the header

    # Try to find the matchup box score
    matchup_box = soup.find('div', id='teamMatchupBoxScore')

    if matchup_box is None:
        # Try alternative selector
        matchup_box = soup.find('div', class_=re.compile('teamMatchup'))

    position_tags = []

    if matchup_box is not None:
        team_wrap = matchup_box.find('div', class_='teamWrap teamWrap-1')

        if team_wrap is None:
            # Try alternative selector
            team_wrap = matchup_box.find('div', class_=re.compile('teamWrap'))

        if team_wrap is not None:
            player_rows = team_wrap.find_all('tr', class_=re.compile('player-'))
            position_tags = [tag.find('span').text for tag in player_rows if tag.find('span')]

    # If we couldn't find position tags, use a default roster structure
    if not position_tags:
        print(f"  Warning: Could not find position tags for week {week}, using default roster")
        # Your league's actual roster structure
        position_tags = ['QB', 'RB', 'RB', 'WR', 'WR', 'TE', 'FLEX', 'K', 'DEF', 'DP',
                         'BN', 'BN', 'BN', 'BN', 'BN', 'BN']

    # position tags are the label for each starting roster spot. different leagues can have different configurations for their starting rosters

    header = []  # csv file header

    # adds the position tags to the header. each tag is followed by a column to record the player's points for the week
    for i in range(len(position_tags)):
        header.append(position_tags[i])
        header.append('Points')

    header = ['Owner', 'Rank'] + header + ['Total', 'Opponent', 'Opponent Total']

    return header


# gets one row of the csv file
# each row is the weekly data for one team in the league
def getrow(teamId, week, longest_bench):
    # loads gamecenter page as soup
    # For 2015, add the gameCenterTab and trackType parameters to load Simple Box Score
    if season == '2015':
        url = f'https://fantasy.nfl.com/league/{leagueID}/history/{season}/teamgamecenter?gameCenterTab=track&teamId={teamId}&trackType=sbs&week={week}'
    else:
        url = f'https://fantasy.nfl.com/league/{leagueID}/history/{season}/teamgamecenter?teamId={teamId}&week={week}'

    page = requests.get(url, cookies=cookies)
    soup = bs(page.text, 'html.parser')
    page.close()

    owner_tag = soup.find('span', class_=re.compile('userName userId'))
    owner = owner_tag.text if owner_tag else f"Team {teamId}"  # username of the team owner

    # Find starters table
    starters_table = soup.find('div', id='tableWrap-1')
    if starters_table is None:
        starters_table = soup.find('div', id=re.compile('tableWrap-[0-9]'))

    if starters_table is not None:
        starters = starters_table.find_all('td', class_='playerNameAndInfo')
        starters = [starter.text for starter in starters]
    else:
        starters = []

    # Handle bench table with error checking
    bench_table = soup.find('div', id='tableWrapBN-1')

    if bench_table is None:
        # Try alternative selector
        bench_table = soup.find('div', id=re.compile('tableWrapBN'))

    if bench_table is not None:
        bench = bench_table.find_all('td', class_='playerNameAndInfo')
        bench = [benchplayer.text for benchplayer in bench]
    else:
        # If no bench found, create empty list
        bench = []

    # in order to keep the row properly aligned, bench spots that are filled by another team
    # but not by this team are filled with a -
    while len(bench) < longest_bench:
        bench.append('-')

    roster = starters + bench  # every player on the team roster, in the order they are listed in game center, for the given week

    # Find player totals with error handling
    matchup_box = soup.find('div', id='teamMatchupBoxScore')
    if matchup_box is None:
        matchup_box = soup.find('div', class_=re.compile('teamMatchup'))

    player_totals = []
    if matchup_box is not None:
        team_wrap = matchup_box.find('div', class_='teamWrap teamWrap-1')
        if team_wrap is None:
            team_wrap = matchup_box.find('div', class_=re.compile('teamWrap'))

        if team_wrap is not None:
            player_totals = team_wrap.find_all('td', class_=re.compile("statTotal"))
            player_totals = [player.text for player in player_totals]

    teamtotals = soup.find_all('div', class_=re.compile('teamTotal teamId-'))  # the team's total points for the week

    rank_tag = soup.find('span', class_=re.compile('teamRank teamId-'))
    if rank_tag:
        ranktext = rank_tag.text
        rank = ranktext[ranktext.index('(') + 1: ranktext.index(')')]  # the team's rank in the standings
    else:
        rank = '-'

    rosterandtotals = []  # alternating player names and their corresponding weekly point totals
    for i in range(len(roster)):
        rosterandtotals.append(roster[i])

        # checks if there is a point total corresponding to the player, if not that spot is filled with a -
        try:
            rosterandtotals.append(player_totals[i])
        except:
            rosterandtotals.append('-')

    # try except statement is for the situation where the league member would not have an opponent for the week
    # in this case the Opponent and Opponent Total columns are filled with -
    try:
        opponent_wrap = soup.find('div', class_='teamWrap teamWrap-2')
        opponent_name = opponent_wrap.find('span', re.compile('userName userId')).text if opponent_wrap else '-'
        opponent_total = teamtotals[1].text if len(teamtotals) > 1 else '-'
        team_total = teamtotals[0].text if len(teamtotals) > 0 else '-'
        completed_row = [owner, rank] + rosterandtotals + [team_total, opponent_name, opponent_total]
    except:
        team_total = teamtotals[0].text if len(teamtotals) > 0 else '-'
        completed_row = [owner, rank] + rosterandtotals + [team_total, '-', '-']

    return completed_row


# Iterate through each season
# Iterate through each week
# Iterate through each team
# Write team's gamecenter data to a csv file

# List of years to skip (already completed successfully)
SKIP_YEARS = []

for s in range(leagueStartYear, leagueEndYear):
    season = str(s)

    # Skip years that are already done
    if s in SKIP_YEARS:
        print(f"Skipping {season} - already completed")
        continue

    # setup
    setup_output_folders(leagueID, season)

    page = requests.get(
        'https://fantasy.nfl.com/league/' + leagueID + '/history/' + season + '/teamgamecenter?teamId=1&week=1',
        cookies=cookies)
    soup = bs(page.text, 'html.parser')
    season_length = len(soup.find_all('li', class_=re.compile(
        'ww ww-')))  # determines how may unique csv files are created, total number of weeks in the season
    number_of_owners = get_number_of_owners(leagueID, season)

    print("Number of Owners: " + str(number_of_owners))
    print("Season Length: " + str(season_length))

    # Iterate through each week of the season, creating a new csv file every loop
    for i in range(1, season_length + 1):
        longest_bench = get_longest_bench(
            i)  # a list containing the length of the longest bench followed by the ID of the team with the longest bench
        header = get_header(i, longest_bench[1])  # header for the csv
        with open(gamecenter_directory + season + '/' + str(i) + '.csv', 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(header)  # writes header as the first line in the new csv file
            for j in range(1, number_of_owners + 1):  # iterates through every team owner
                writer.writerow(getrow(str(j), str(i), longest_bench[0]))  # writes a row for each owner in the csv
        print("Week " + str(i) + " Complete")
    print("Done")