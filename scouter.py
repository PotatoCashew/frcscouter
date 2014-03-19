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
for team in get_teams():
    sheet = workbook.add_sheet(str(team[u'team_number']))
    for (i, item) in enumerate(team):
        sheet.write(0, i, item)
        sheet.write(1, i, team[item])

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