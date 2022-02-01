from asyncio.windows_events import NULL
import csv
import arrow
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import requests
from dotenv import load_dotenv

load_dotenv()

clientsFilename = 'clients.csv'
filterOutListFilename = 'filter.csv'
visitsFilename = ''

def requestSchedule(clients, filterOutList):
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
    accessToken = tokenRes['access_token']
    scheduleParams = {
        'access_token': accessToken,
        'start_date': startDate.get(),
        'end_date': endDate.get(),
        'fields': 'employee,location,start_day,end_day,start_time,end_time,total_time',
        'type': 'shifts'
    }
    scheduleReq = requests.get('https://www.humanity.com/api/v2/reports/custom?', params=scheduleParams)
    scheduleRes = scheduleReq.json()
    scheduleReport = scheduleRes['data']
    schedule = []

    for key in scheduleReport:
        reportRow = scheduleReport[key]
        if key.isnumeric():
            if reportRow['location'] != '':
                reportRow['location'] = flipName(reportRow['location'])
                if reportRow['location'] in clients and reportRow['employee'] not in filterOutList:
                    reportRow['employee'] = flipName(reportRow['employee'])
                    schedule.append(reportRow)
    return schedule

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

# Open clients csv file and return list of clients
def getClientsFromFile():
    clients = []
    if os.path.exists(clientsFilename):
        with open(clientsFilename, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                clients.append(row['CLIENT LAST NAME'] + ', ' + row['CLIENT FIRST NAME'])
    else:
        messagebox.showwarning(message='visits.csv file is empty or does not exist!')
    return clients

# Open visits csv file
# return visits as a list
def getVisitsFromFile():
    visits = []
    if os.path.exists(visitsFilename):
        with open(visitsFilename, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                visits.append(row)
    return visits

# Write the contents of paramter missing list to MissingShifts.csv file
# Changes keys (header) to make the most sense to the end user
def writeMissingShiftsFile(missing):
    with open('MissingShifts.csv', 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['Client', 'Scheduled Staff', 'Date', 'Start', 'End'])
        writer.writeheader()
        for missed in missing:
            writer.writerow({'Client': missed['location'], 'Scheduled Staff': missed['employee'], 'Date': missed['start_day'], 'Start': missed['start_time'], 'End': missed['end_time']})

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

def getFilterList():
    filterOutList = []
    if os.path.exists(filterOutListFilename):
        with open(filterOutListFilename, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                filterOutList.append(row['filterOut'])
    return filterOutList

# Retrieve client list, schedule, and vists from files or return if not selected/empty
# For each record in schedule, if a visit is not found that matches it, add it to missing list
# Write missing list to a missingShifts.csv
def findMissingVisits():
    missing = []
    filterOutList = getFilterList()
    clients = getClientsFromFile()
    schedule = requestSchedule(clients, filterOutList)
    # schedule = getScheduleFromFile(clients, filterOutList)
    visits = getVisitsFromFile()
    
    if not schedule:
        messagebox.showwarning(message='No schedule found.')
        return
        
    if not visits:
        if not visitsFilename:
            messagebox.showwarning(message='No visits file selected. Please select a visits file first.')
        else:
            messagebox.showwarning(message=f'Could not find any visits in visits file: {visitsFilename}')
        return
    
    for record in schedule:
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

    messagebox.showinfo(message=str(round(len(missing) / len(schedule) * 100)) + '% of shifts missing (' + str(len(missing)) + '/' + str(len(schedule)) + ')')
    writeMissingShiftsFile(missing)
    

# GUI
root = Tk() 
root.title('Check Missing Shifts')

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

startDate = StringVar()
endDate = StringVar()
label = ttk.Label(mainframe, text='Find missing shifts from a schedule and a visits file.\nSchedule file: Exported .csv file from the custom report called "Scheduled Shifts" in Humanity Scheduler\nVisits file: Exported .csv file from Sandata EVV with filtered "visit type" as "verified"\nSchedule and visits files should be for the same date range\nFind missing shifts will create a "missingShifts.csv" file').grid(column=1, row=1)
startDateLabel = ttk.Label(mainframe, text="Start Date").grid(column=1, row=2, sticky=(E))
startDateEntry = ttk.Entry(mainframe, textvariable=startDate).grid(column=2, row=2,  sticky=(W))
endDateLabel = ttk.Label(mainframe, text="End Date").grid(column=1, row=3, sticky=(E))
endDateEntry = ttk.Entry(mainframe, textvariable=endDate).grid(column=2, row=3,  sticky=(W))
openVisitsButton = ttk.Button(mainframe, text='Select Visits File', command=openVisitsDialog).grid(column=1, row=4, sticky=(W))
findMissingButton = ttk.Button(mainframe, text='Find Missing Visits', command=findMissingVisits).grid(column=1, row=5)

root.mainloop()