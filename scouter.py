import urllib2, xlwt, simplejson as json, tkFileDialog
from os.path import expanduser

url = "http://www.thebluealliance.com/api/v2/event/"
HEADER_NAME = "x-tba-app-id"
header649 = "frc649:scouter:1"

print("year:")
year = raw_input()
print("event code (e.g. cama)")
event_code = raw_input()
url += year + event_code + '/'

def get_teams():
    req = urllib2.Request(url + "teams")
    req.add_header(HEADER_NAME, header649)
    response = urllib2.urlopen(req)
    return json.loads(response.read())

workbook = xlwt.Workbook()

sheet = workbook.add_sheet(year + event_code)
i = 0
sheet.write(0, i, "Number")
i += 1
sheet.write(0, i, "Nickname")
i += 1
sheet.write(0, i, "1st Regional")
i += 1
sheet.write(0, i, "Record")
i += 1
sheet.write(0, i, "OPR")
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

for (index, team) in enumerate(get_teams()):
    sheet.write(index+1, 0, team[u'team_number'])
    sheet.write(index+1, 1, team[u'nickname'])

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