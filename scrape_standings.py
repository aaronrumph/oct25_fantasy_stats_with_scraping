import csv
from bs4 import BeautifulSoup as bs
import requests
from cookie_string import cookies
from utils import setup_output_folders
from constants import leagueID, leagueStartYear, leagueEndYear, standings_directory

# iterate through each season
# parse standings, owners, and draft results
# write to a csv file
for i in range(leagueStartYear, leagueEndYear):
    season = str(i)
    setup_output_folders(leagueID, season)

    # parse Regular Season Standings
    # https://fantasy.nfl.com/league/1609009/history/2023/standings?historyStandingsType=regular
    page = requests.get(
        'https://fantasy.nfl.com/league/' + leagueID + '/history/' + season + '/standings?historyStandingsType=regular',
        cookies=cookies)
    soup = bs(page.text, 'html.parser')
    csv_rows = []

    # parse the regular season standings table
    # adds cols: 'TeamName', 'RegularSeasonRank', 'Record', 'PointsFor', 'PointsAgainst'
    season_table_rows = soup.find_all('tr', class_=lambda x: x and 'team' in x)
    for row in season_table_rows:
        season_rank = row.find('span', class_='teamRank').text.strip()
        team_name = row.find('a', class_='teamName').text.strip()
        team_record = row.find('td', class_='teamRecord').text.strip()
        pts_for = row.find_all('td', class_='teamPts')[0].text.strip()
        pts_against = row.find_all('td', class_='teamPts')[1].text.strip()
        csv_rows.append([team_name, season_rank, team_record, pts_for, pts_against])

    # parse Playoffs Season Standings
    # https://fantasy.nfl.com/league/1609009/history/2023/standings?historyStandingsType=final
    page = requests.get(
        'https://fantasy.nfl.com/league/' + leagueID + '/history/' + season + '/standings?historyStandingsType=final',
        cookies=cookies)
    soup = bs(page.text, 'html.parser')

    # parse the playoffs standings table
    # adds col: 'PlayoffRank'
    playoff_table_rows = soup.find_all('li', class_=lambda x: x and 'place' in x)
    for row in playoff_table_rows:
        # find the div with class 'place' to extract the place number
        place_div = row.find('div', class_='place')
        if place_div:
            # extract the place number from the text
            place_number = place_div.text.split()[0][:-2]
            # find the anchor tag within the div with class 'value' to extract the team name
            team_name_anchor = row.find('div', class_='value').find('a', class_='teamName')
            if team_name_anchor:
                team_name = team_name_anchor.text.strip()
                for csv_row in csv_rows:
                    if csv_row[0] == team_name:
                        csv_row.append(place_number)

    # parse Owners
    # https://fantasy.nfl.com/league/1609009/history/2023/owners
    page = requests.get('https://fantasy.nfl.com/league/' + leagueID + '/history/' + season + '/owners',
                        cookies=cookies)
    soup = bs(page.text, 'html.parser')

    # parse the owners table
    # adds cols: 'ManagerName', 'Moves', 'Trades'
    season_table_rows = soup.find_all('tr', class_=lambda x: x and 'team' in x)
    for row in season_table_rows:
        team_name = row.find('a', class_='teamName').text.strip()
        manager = row.find('span', class_='userName').text.strip()
        moves = row.find('td', class_='teamTransactionCount').text.strip()
        trades = row.find('td', class_='teamTradeCount').text.strip()
        for csv_row in csv_rows:
            if csv_row[0] == team_name:
                csv_row.append(manager)
                csv_row.append(moves)
                csv_row.append(trades)

    # parse Draft Results
    # https://fantasy.nfl.com/league/1609009/history/2023/draftresults
    page = requests.get('https://fantasy.nfl.com/league/' + leagueID + '/history/' + season + '/draftresults',
                        cookies=cookies)
    soup = bs(page.text, 'html.parser')

    # parse the Draft results table
    # adds cols: 'DraftPosition'
    round_1_ul = soup.find('h4', string='Round 1').find_next_sibling('ul')
    round_1_li_elements = round_1_ul.find_all('li')
    for row in round_1_li_elements:
        draft_position = row.find('span', class_='count')
        team_name = row.find('a', class_='teamName')
        if draft_position and team_name:
            for csv_row in csv_rows:
                if csv_row[0] == team_name.text.strip():
                    csv_row.append(draft_position.text.strip()[:-1])

    # write all to a csv file
    with open(standings_directory + season + '.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        header = ['TeamName', 'RegularSeasonRank', 'Record', 'PointsFor', 'PointsAgainst', 'PlayoffRank', 'ManagerName',
                  'Moves', "Trades", "DraftPosition"]
        writer.writerow(header)
        for row in csv_rows:
            writer.writerow(row)

    print(season + " parsed.")

