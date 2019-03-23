from icalendar import Calendar, Event
from datetime import datetime
from stat import S_ISREG, ST_CTIME, ST_MODE
import os, sys, time


dir_path = 'D:\\Firefox Downloads\\'

#all entries in the directory w/ stats
data = (os.path.join(dir_path, fn) for fn in os.listdir(dir_path) if fn.endswith(".ics"))
data = ((os.stat(path), path) for path in data)

# regular files, insert creation date
data = ((stat[ST_CTIME], path)
           for stat, path in data if S_ISREG(stat[ST_MODE]))

file = sorted(data, reverse=True)[0][1]

try:
    g = open(file,'rb')
    o = open('C:\\Users\\Alice\\Desktop\\birthdays.txt', 'wb')
except IOError:
    print "Could not open file! Please close Excel!"

gcal = Calendar.from_ical(g.read())
current_date = datetime.now().__str__().split()[0]
i = 0
people = []
for component in gcal.walk():
    if component.name == "VEVENT":
    	persone_date = component.get('dtstart').__dict__['dt'].__str__()
    	if persone_date == current_date:
    		uid = component.get('uid').split('@')[0]
    		uid = uid[1:]
    		link = 'https://www.facebook.com/profile.php?id='
    		link = link + uid
    		name = component.get('summary').split()[0]
    		simbol = ':*' if name[len(name) - 1] == 'a' else ':)'
        	lma = 'La multi ani, ' + name  + '! ' + simbol
        	people.append((link, lma))
        else:
        	break
o.write(str(people.__len__()) + "\n")
for (link, lma) in people:
	o.write(link + "\n" + lma + "\n")
g.close()