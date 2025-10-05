import os
from bs4 import BeautifulSoup as bs
from urllib.request import urlopen
import re
import requests
from cookie_string import cookies

#gets the total number of players in a given season
def get_number_of_owners(leagueID, season) :
	owners_url = 'https://fantasy.nfl.com/league/' + leagueID + '/history/' + season + '/owners'
	owners_page = requests.get(owners_url, cookies=cookies)
	owners_html = owners_page.text
	owners_soup = bs(owners_html, 'html.parser')
	number_of_owners = len(owners_soup.find_all('tr', class_ = re.compile('team-')))
	return number_of_owners

def setup_output_folders(leagueID, season):
    path = "./output/"
    if not os.path.isdir(path):
        os.mkdir(path)

    games_path = path + leagueID + "-history-teamgamecenter/"
    if not os.path.isdir(games_path):
        os.mkdir(games_path)

    games_path += season + "/"
    if not os.path.isdir(games_path):
        os.mkdir(games_path)

    standings_path = "./output/" + leagueID + "-history-standings/"
    if not os.path.isdir(standings_path):
	    os.mkdir(standings_path)
