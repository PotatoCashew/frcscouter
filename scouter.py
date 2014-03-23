import urllib2, xlwt, simplejson as json, tkFileDialog, time, re, os
from os.path import expanduser
from operator import itemgetter
from selenium import webdriver

url = "http://www.thebluealliance.com/api/v2/"
HEADER_NAME = "x-tba-app-id"
header649 = "frc649:scouter:1"
adambots_headers = ['Rank', 'Auton', 'Truss/Catch', 'Assist', 'Other Teleop', 'OPR (Total)', 'CCWM', 'Predicted QS']
print("year:")
year = raw_input()
first_alliance_url = "http://www.thefirstalliance.org/api/api.json.php?action="
print("event code (e.g. cama):")
event_code = raw_input()

championships_url = 'https://my.usfirst.org/myarea/index.lasso?page=teamlist&event_type=FRC&sort_teams=number&year=' + year + '&event=cmp'
regional_data_url = 'http://www2.usfirst.org/'+year+'comp/Events/$1/rankings.html'
championships_regex = '<tr bgcolor="#.*?">(\s*<td.*?/td>\s+){2}<td><a href.*?>([0-9]+)'
regional_data_team_regex = '<TR.*?>(?:\s*.*>.*</TD>)+\s*</TR>'
regional_data_data_regex = '<TD.*>(.*)<'

browser = webdriver.Chrome()

def send_json_request(url):
    req = urllib2.Request(url)
    req.add_header(HEADER_NAME, header649)
    response = urllib2.urlopen(req)
    return json.loads(response.read())
    
def send_plaintext_request(url):
    req = urllib2.Request(url)
    req.add_header(HEADER_NAME, header649)
    response = urllib2.urlopen(req)
    return response.read()
    
def get_teams():
    return send_json_request(url + "event/" + year + event_code + '/' + "teams")

def get_detailed_team_info(team_number):
    return send_json_request(url + "team/frc" + str(team_number) + "/" + year)
    
def write_headers(num_regionals):    
    i = 0
    sheet.write(0, i, "Number")
    i += 1
    sheet.write(0, i, "Nickname")
    i += 1
    sheet.write(0, i, "Going to Champs")
    i += 1
    for x in range(num_regionals):
        sheet.write(0, i, "Regional " + str(x+1))
        i += 1
        sheet.write(0, i, "Record")
        i += 1
        sheet.write(0, i, "Rank")
        i += 1
        sheet.write(0, i, "OPR")
        i += 1
        sheet.write(0, i, "DPR")
        i += 1
        sheet.write(0, i, "CCWM")
        i += 1
        sheet.write(0, i, "Assist Points")
        i += 1
        sheet.write(0, i, "Auto Points")
        i += 1
        sheet.write(0, i, "T&C Points")
        i += 1
        sheet.write(0, i, "Teleop Points")
        i += 1
        sheet.write(0, i, "DQ")
        i += 1
        sheet.write(0, i, "Awards")
        i += 1
    
def get_regional_advanced_stats(regional):
    url = "http://www.adambots.com/scouting/automated2014/?comp="
    req = urllib2.Request(url + regional)
    browser.get(url + regional)
    while "(Analysis)" not in browser.page_source:
        time.sleep(0.01)
        if "Not enough matches have been played." in browser.page_source:
            return None
    regional = {}
    for row in browser.find_element_by_id("bigtableplace").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr"):
        team = []
        for (i, td) in enumerate(row.find_elements_by_tag_name("td")):
            if i == 0:
                team_number = td.text
            else:
                team.append(td.text)
        regional[int(team_number)] = team
    return regional
def get_championships_teams():
    text = send_plaintext_request(championships_url)
    matches = re.findall(championships_regex, text)
    teams = []
    for match in matches:
        teams.append(match[1])
    return teams  
def aggregate_regional_results(regional_code):
    try:
        text = send_plaintext_request(regional_data_url.replace('$1', regional_code))
    except:
        return None
    teams = re.findall(regional_data_team_regex, text)
    regional_data = {}
    for team in teams:
        team_stats = team.split('/TD>')
        team_number = re.findall(regional_data_data_regex, team_stats.pop(1))
        team_stats.pop()
        team_number = team_number[0]
        team_regional_stats = []
        for team_stat in team_stats:
            team_regional_stats.append(re.findall(regional_data_data_regex, team_stat)[0])
        regional_data[int(team_number)] = team_regional_stats
    return regional_data
        
championship_team_list = get_championships_teams()
workbook = xlwt.Workbook()

sheet = workbook.add_sheet(year + event_code)
max_regionals = 0
teams = sorted(get_teams(), key=itemgetter('team_number'))
regional_advanced_data = {}
regional_stats_data = {}
team_infos = []

for (index, team) in enumerate(teams):
    team_number = team[u'team_number']
    team_nickname = team[u'nickname']
    sheet.write(index+1, 0, team_number)
    sheet.write(index+1, 1, team_nickname)
    team_info = get_detailed_team_info(team_number)
    if str(team_number) in championship_team_list:
        sheet.write(index+1, 2, "Yes")
    else:
        sheet.write(index+1, 2, "No")
    print("Getting data on " + str(team_number) + ": " + team_nickname + " (" + str(index + 1) + "/" + str(len(teams)) + ")")
    max_regionals = max(len(team_info[u'events'])-1, max_regionals)
    if len(team_info[u'events']) == 0:
        sheet.write(index+1, 3,"None")
    for (regional_index, regional) in enumerate(team_info[u'events']):
        regional_event_code = regional[u'event_code']
        regional_name = regional[u'name']
        if not regional_event_code == event_code:
            if not regional_event_code in regional_advanced_data:
                print('Collecting data on ' + regional_name)
                regional_advanced_data[regional_event_code] = get_regional_advanced_stats(regional_name.replace(' ', '-'))
                regional_stats_data[regional_event_code] = aggregate_regional_results(regional_event_code)
            sheet.write(index+1, 3 + 9*regional_index,regional_name)
            if regional_advanced_data[regional_event_code] is not None:        
                team_stats_data = regional_stats_data[regional_event_code][team_number]
                sheet.write(index+1, 4 + 9*regional_index,team_stats_data[6])
                sheet.write(index+1, 5 + 9*regional_index,float(team_stats_data[0]))
                opr = regional_advanced_data[regional_event_code][team_number][5]
                ccwm = regional_advanced_data[regional_event_code][team_number][6]
                sheet.write(index+1, 6 + 9*regional_index, str(opr))
                sheet.write(index+1, 7 + 9*regional_index, str(float(opr)-float(ccwm)))
                sheet.write(index+1, 8 + 9*regional_index, str(ccwm))
                sheet.write(index+1, 9 + 9*regional_index,float(team_stats_data[2]))
                sheet.write(index+1, 10 + 9*regional_index,float(team_stats_data[3]))
                sheet.write(index+1, 11 + 9*regional_index,float(team_stats_data[4]))
                sheet.write(index+1, 12 + 9*regional_index,float(team_stats_data[5]))
                sheet.write(index+1, 13 + 9*regional_index,float(team_stats_data[7]))
                awards = ""
                for award in regional[u'awards']:
                    awards = awards + award[u'name'] + ", "
                if len(awards) > 0:
                    awards = awards[:-2]
                sheet.write(index+1, 14 + 9 *regional_index, awards)

write_headers(max_regionals)

options = {}
options['defaultextension'] = '.xls'
options['filetypes'] = [('all files', '.*'), ('Excel spreadsheets', '.xls')]
options['initialdir'] = expanduser("~")
options['initialfile'] = year + event_code + ".xls"
options['title'] = 'Save as'
try:
    filename = tkFileDialog.asksaveasfilename(**options)
    try:
        os.remove(filename)
    except OSError:
        pass
    workbook.save(filename)
except:
    print("No file specified")
    raise
    
browser.close()