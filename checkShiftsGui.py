from asyncio.windows_events import NULL
import csv
import arrow
from os.path import exists
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

clientsFilename = 'clients.csv'
filterOutListFilename = 'filter.csv'
scheduleFilename = ''
visitsFilename = ''

# Open a choose file dialog for schedule
# Called from button in GUI
def openScheduleDialog():
    global scheduleFilename
    scheduleFilename = filedialog.askopenfilename()

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
    if exists(clientsFilename):
        with open(clientsFilename, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                clients.append(row['CLIENT LAST NAME'] + ', ' + row['CLIENT FIRST NAME'])
    return clients

# Open schedule csv file
# return list of schedules that have a client in paramter clients list
# and an employee that does not exists in paramter filterOutList
def getScheduleFromFile(clients, filterOutList):
    schedule = [] 
    if exists(scheduleFilename):    
        with open(scheduleFilename, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['location'] != '':
                    row['location'] = flipName(row['location'])
                    if row['location'] in clients and row['employee'] not in filterOutList:
                        row['employee'] = flipName(row['employee'])
                        schedule.append(row)
    return schedule

# Open visits csv file
# return visits as a list
def getVisitsFromFile():
    visits = []
    if exists(visitsFilename):
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
    if exists(filterOutListFilename):
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
    schedule = getScheduleFromFile(clients, filterOutList)
    visits = getVisitsFromFile()
    
    if not clients:
        messagebox.showwarning(message='visits.csv file is empty or does not exist! Aborting..')
        return
    if not schedule:
        if not scheduleFilename:
            messagebox.showwarning(message='No schedule file selected. Please select a schedule file first.')
        else:
            messagebox.showwarning(message=f'Could not find any schedules for clients in Schedule file: {scheduleFilename}')
        return

    if not visits:
        if not visitsFilename:
            messagebox.showwarning(message='No visits file selected. Please select a visits file first.')
        else:
            messagebox.showwarning(message=f'Could not find any visits in visits file: {visitsFilename}')
        return
    
    for record in schedule:
        match = False
        scheduledStart = arrow.get(record['start_day'] + ' ' + record['start_time'], 'MM/DD/YYYY h:mmA')
        scheduledEnd = arrow.get(record['end_day'] + ' ' + record['end_time'], 'MM/DD/YYYY h:mmA')
        for visit in visits:
            if visit['Client Name'] == record['location']:
                visitStart, visitEnd = getVisitDatetimes(visit)

                if timeWithinSpan(visitStart, scheduledStart, 30) and timeWithinSpan(visitEnd, scheduledEnd, 30):
                    match = True
                    break
        
        if match == False:
            missing.append(record)

    print('Missing ' + str(len(missing)) + ' out of ' + str(len(schedule)) + ' scheduled and ' + str(len(visits)) + ' visits')
    writeMissingShiftsFile(missing)
    

# GUI
root = Tk() 
root.title('Check Missing Shifts')

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

label = ttk.Label(mainframe, text='Find missing shifts from a schedule and a visits file.\nSchedule file: Exported .csv file from the custom report called "Scheduled Shifts" in Humanity Scheduler\nVisits file: Exported .csv file from Sandata EVV with filtered "visit type" as "verified"\nSchedule and visits files should be for the same date range\nFind missing shifts will create a "missingShifts.csv" file').grid(column=1, row=1)
openScheduleButton = ttk.Button(mainframe, text='Select Schedule File', command=openScheduleDialog).grid(column=1, row=2, sticky=(W))
openVisitsButton = ttk.Button(mainframe, text='Select Visits File', command=openVisitsDialog).grid(column=1, row=3, sticky=(W))
findMissingButton = ttk.Button(mainframe, text='Find Missing Visits', command=findMissingVisits).grid(column=1, row=5)

root.mainloop()