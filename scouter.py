import urllib2, xlwt, simplejson as json, tkFileDialog, time
from os.path import expanduser
from operator import itemgetter
from selenium import webdriver

url = "http://www.thebluealliance.com/api/v2/"
HEADER_NAME = "x-tba-app-id"
header649 = "frc649:scouter:1"

print("year:")
year = raw_input()
first_alliance_url = "http://www.thefirstalliance.org/api/api.json.php?action="
print("event code (e.g. cama)")
event_code = raw_input()

browser = webdriver.Chrome()

def send_request(url):
    req = urllib2.Request(url)
    req.add_header(HEADER_NAME, header649)
    response = urllib2.urlopen(req)
    return json.loads(response.read())
    
def get_teams():
    return send_request(url + "event/" + year + event_code + '/' + "teams")

def get_detailed_team_info(team_number):
    return send_request(url + "team/frc" + str(team_number) + "/" + year)
    
def write_headers(num_regionals):    
    i = 0
    sheet.write(0, i, "Number")
    i += 1
    sheet.write(0, i, "Nickname")
    i += 1
    for x in range(num_regionals):
        sheet.write(0, i, "Regional " + str(x+1))
        i += 1
        sheet.write(0, i, "Record")
        i += 1
        sheet.write(0, i, "power_rating")
        i += 1
        sheet.write(0, i, "Auto Points")
        i += 1
        sheet.write(0, i, "Assist Points")
        i += 1
        sheet.write(0, i, "T&C Points")
        i += 1
        sheet.write(0, i, "Teleop Points")
        i += 1
        sheet.write(0, i, "DPR")
        i += 1
    sheet.write(0, i, "Going to Champs")
    i += 1
    sheet.write(0, i, "Awards")
    
def get_regional_advanced_stats(regional):
    url = "http://www.adambots.com/scouting/automated2014/?comp="
    req = urllib2.Request(url + regional)
    browser.get(url + regional)
    while "(Analysis)" not in browser.page_source:
        time.sleep(0.01)
    regional = {}
    for row in browser.find_element_by_id("bigtableplace").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr"):
        team = {}
        first = True
        for td in row.find_elements_by_tag_name("td"):
            if first:
                first = False
                team_number = id.text
            else:
                team[td.text
    return regional
    
workbook = xlwt.Workbook()

sheet = workbook.add_sheet(year + event_code)
max_regionals = 0
teams = sorted(get_teams(), key=itemgetter('team_number'))
regional_data = {}
for (index, team) in enumerate(teams):
    team_number = team[u'team_number']
    team_nickname = team[u'nickname']
    sheet.write(index+1, 0, team_number)
    sheet.write(index+1, 1, team_nickname)
    team_info = get_detailed_team_info(team_number)
    print("Getting data on " + str(team_number) + ": " + team_nickname + " (" + str(index + 1) + "/" + str(len(teams)) + ")")
    max_regionals = len(team_info[u'events'])-1
    for regional in team_info[u'events']:
        regional_event_code = regional[u'event_code']
        regional_name = regional[u'name'].replace(' ', '-')
        if not regional_event_code == event_code:
            print(regional_name)
            print(get_regional_advanced_stats(regional_name))
        #if not regional[u'event_code'] in regional_data:
            #regional_data[regional[u'event_code']] = 

write_headers(max_regionals)

options = {}
options['defaultextension'] = '.xls'
options['filetypes'] = [('all files', '.*'), ('Excel spreadsheets', '.xls')]
options['initialdir'] = expanduser("~")
options['initialfile'] = year + event_code + ".xls"
options['title'] = 'Save as'
try:
    workbook.save(tkFileDialog.asksaveasfilename( **options))
except:
    print("No file specified")
    
browser.close()