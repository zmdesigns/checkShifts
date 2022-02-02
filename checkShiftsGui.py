import csv
from json.encoder import INFINITY
from sysconfig import get_scheme_names
import arrow
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import requests
from dotenv import load_dotenv
from datetime import date

load_dotenv()
visitsFilename = ''

class Schedule:
    '''A class to hold a schedule requested from scheduler api'''

    def __init__(self, startDate, endDate, clients, filterEmployees):
        self.startDate = startDate
        self.endDate = endDate
        self.clients = clients
        self.filterEmployees = filterEmployees
        self.schedule = []

    def _getToken(self):
        userDetails = {
            'client_id': os.environ["CHECKSHIFTS_CLIENTID"],
            'client_secret': os.environ["CHECKSHIFTS_CLIENTSECRET"],
            'grant_type': 'password',
            'username': os.environ["CHECKSHIFTS_USERNAME"],
            'password': os.environ["CHECKSHIFTS_PASSWORD"],
            'redirect_uri': ''
        }
        tokenReq = requests.post('https://www.humanity.com/oauth2/token.php', data=userDetails)
        tokenRes = tokenReq.json()
        return tokenRes['access_token']

    def _getReport(self, accessToken):
        scheduleParams = {
            'access_token': accessToken,
            'start_date': self.startDate,
            'end_date': self.endDate,
            'fields': 'employee,location,start_day,end_day,start_time,end_time,total_time',
            'type': 'shifts'
        }
        print('Requesting scheduled shifts from ' + scheduleParams['start_date'] + ' to ' + scheduleParams['end_date'] + ' from scheduler API')
        scheduleReq = requests.get('https://www.humanity.com/api/v2/reports/custom?', params=scheduleParams)
        scheduleRes = scheduleReq.json()
        return scheduleRes['data']

    def getSchedule(self):
        accessToken = self._getToken()
        report = self._getReport(accessToken)
        for key in report:
            reportRow = report[key]
            if key.isnumeric():
                if reportRow['location'] != '':
                    reportRow['location'] = flipName(reportRow['location'])
                    if reportRow['location'] in self.clients and reportRow['employee'] not in self.filterEmployees:
                        reportRow['employee'] = flipName(reportRow['employee'])
                        self.schedule.append(reportRow)

# Can't assign instance values to results of functions not declared yet
class CheckShifts:
    '''A class to check for missing shifts'''

    def __init__(self, visitsFilename):
        self.startDate = arrow.get(date.max)
        self.endDate = arrow.get(date.min)
        self.visitsFilename = visitsFilename

    def _getFilterList(self):
        filterOutList = []
        if os.path.exists('filter.csv'):
            with open('filter.csv', newline='') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    filterOutList.append(row['filterOut'])
        return filterOutList

    # Open visits csv file
    # return visits as a list
    def _getVisitsAndClientsFromFile(self, visitsFilename):
        visits = []
        clients = set()
        if os.path.exists(visitsFilename):
            with open(visitsFilename, newline='') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    clients.add(row['Client Name'])
                    visits.append(row)
                    visitDate = arrow.get(row['Visit Date'], 'MM/DD/YYYY');
                    if visitDate < self.startDate:
                        self.startDate = visitDate
                    if visitDate > self.endDate:
                        self.endDate = visitDate
        print('Identified ' + str(len(visits)) + ' visits for ' + str(len(clients)) + ' unique clients between ' + self.startDate.format('MM/DD/YY') + ' - ' + self.endDate.format('MM/DD/YY'))
        return visits, clients

    # Write the contents of paramter missing list to MissingShifts.csv file
    # Changes keys (header) to make the most sense to the end user
    def _writeMissingShiftsFile(self, missing):
        with open('MissingShifts.csv', 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=['Client', 'Scheduled Staff', 'Date', 'Start', 'End'])
            writer.writeheader()
            for missed in missing:
                writer.writerow({'Client': missed['location'], 'Scheduled Staff': missed['employee'], 'Date': missed['start_day'], 'Start': missed['start_time'], 'End': missed['end_time']})

    # Retrieve client list, schedule, and vists from files or return if not selected/empty
    # For each record in schedule, if a visit is not found that matches it, add it to missing list
    # Write missing list to a missingShifts.csv
    def findMissingVisits(self):
        visits, clients = self._getVisitsAndClientsFromFile(self.visitsFilename)
        filterEmployees = self._getFilterList()
        missing = []

        schedule = Schedule(self.startDate.format('YYYY/MM/DD'), self.endDate.format('YYYY/MM/DD'), clients, filterEmployees)
        schedule.getSchedule()
        
        if not schedule.schedule:
            messagebox.showwarning(message='No schedule found.')
            return
            
        if not visits:
            messagebox.showwarning(message='Could not find any visits in visits file')
            return
        
        for record in schedule.schedule:
            match = False
            scheduledStart = arrow.get(record['start_day'] + ' ' + record['start_time'], 'MMM D, YYYY h:mmA')
            scheduledEnd = arrow.get(record['end_day'] + ' ' + record['end_time'], 'MMM D, YYYY h:mmA')
            for visit in visits:
                if visit['Client Name'] == record['location']:
                    visitStart, visitEnd = getVisitDatetimes(visit)

                    if timeWithinSpan(visitStart, scheduledStart, 30) and timeWithinSpan(visitEnd, scheduledEnd, 30):
                        match = True
                        break
            
            if match == False:
                missing.append(record)

        messagebox.showinfo('Done')
        self._writeMissingShiftsFile(missing)

# Open a choose file dialog for visits
# Called from button in GUI
def openVisitsDialog():
    global visitsFilename
    visitsFilename = filedialog.askopenfilename()

# Flips First Last formated name to Last, First
def flipName(name):
    nameList = name.split()
    if (len(nameList) > 1):
        return nameList[1] + ', ' + nameList[0]
    return name

# Return if visit time is within range of scheduled time +/- span minutes
def timeWithinSpan(timeToCheck, spanMiddle, spanRadius):
    spanStart = spanMiddle.shift(minutes=-spanRadius)
    spanEnd = spanMiddle.shift(minutes=+spanRadius)

    if timeToCheck > spanStart and timeToCheck < spanEnd:
        return True
    return False

# Return start and end of visit given visit row
# computes visit end by finding what the visit length is
# (depending on if the visit was adjusted, it could be "Call Hours" or "Adjusted Hours")
# Shifts the start date by the visit length (Hours:Minutes) to find visit end
# This is needed because "Visit Date" is the date when the employee started the visit
# Using "Visit Date" + "Adjusted Out" can give the wrong datetime because of this
# ie- Visit Date: 1/2/22 Adjusted In: 11:59PM Adjusted Out: 8:00AM
def getVisitDatetimes(visit):
    visitStart = arrow.get(visit['Visit Date'] + ' ' + visit['Adjusted In'], 'MM/DD/YYYY h:mm A')
    callOrAdj = 'Call Hours' if visit['Call Hours'].strip() else 'Adjusted Hours'
    visitLength = visit[callOrAdj].split(':')

    if len(visitLength) <= 1:
        print(visitLength)
        return NULL

    visitEnd = visitStart.shift(hours=int(visitLength[0]),minutes=int(visitLength[1]))
    return visitStart, visitEnd

def checkVisits():
    missingVisits = CheckShifts(visitsFilename)
    missingVisits.findMissingVisits()

# GUI
root = Tk() 
root.title('Check Missing Shifts')

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

label = ttk.Label(mainframe, text='Find missing shifts from a schedule and a visits file.\nSchedule file: Exported .csv file from the custom report called "Scheduled Shifts" in Humanity Scheduler\nVisits file: Exported .csv file from Sandata EVV with filtered "visit type" as "verified"\nSchedule and visits files should be for the same date range\nFind missing shifts will create a "missingShifts.csv" file').grid(column=1, row=1)
openVisitsButton = ttk.Button(mainframe, text='Select Visits File', command=openVisitsDialog).grid(column=1, row=4, sticky=(W))
findMissingButton = ttk.Button(mainframe, text='Find Missing Visits', command=checkVisits).grid(column=1, row=5)

root.mainloop()