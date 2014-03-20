import urllib2, xlwt, simplejson as json, tkFileDialog, urllib
from os.path import expanduser
from operator import itemgetter

url = "http://www.thebluealliance.com/api/v2/"
first_alliance_url = "http://www.thefirstalliance.org/api/api.json.php?action="
HEADER_NAME = "x-tba-app-id"
header649 = "frc649:scouter:1"

print("year:")
year = raw_input()
print("event code (e.g. cama)")
event_code = raw_input()

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
	
def get_regional_data(event_code):
	return None
	
def get_power_rating(team, regional, type):
    url = "http://www.adambots.com/scouting/automated2014/?comp="
    #parameters = {"team-number" : str(team), "event-code" : regional}
    #data = urllib.urlencode(parameters)
    #req = urllib2.Request(url + "team-event-" + type.lower(), data)
    req = urllib2.Request(url + regional)
    response1 = urllib2.urlopen(req)
    print(response1.read())
    #power_rating_data = json.loads(response1.read())
    index = 0
    power_rating = None
    while (power_rating is None or power_rating is not "null") and index < len(power_rating_data[u'data']):
		power_rating = power_rating_data[u'data'][index]['OPR' if type.lower() == 'opr' else 'dPR']
		index+= 1
	
    return (power_rating)
	
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
		print(regional)
		team_event_code = regional[u'event_code']
		if not team_event_code == event_code:
			print(get_power_rating(team_number, team_event_code, 'opr'), get_power_rating(team_number, team_event_code, 'dpr'))
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
    quit(0)
quit(0)