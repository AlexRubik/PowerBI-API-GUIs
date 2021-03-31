'''
GUI for PowerBI Dataflow APIs (not all of them)
Contact: dewittat@g.cofc.edu
'''


from tkinter import *
from tkinter import ttk
import Dataflow_APIs
import webbrowser



root = Tk()
root.title("PowerBI Dataflows GUI")
root.geometry("870x790+450+100")
heading = Label(root, text = "PowerBI Dataflows GUI", font = ("arial",25,"bold"),fg = "purple").pack()

#------------------------------------------------------------------------------------------------------------------------------------------------------
# Email/Password Labels/Entry Boxes
#------------------------------------------------------------------------------------------------------------------------------------------------------

emailLabel = Label(root, text = "Microsoft Email:", font = ("arial", 10, "bold"), fg = "black").place(x=5,y=165)

passwordLabel = Label(root, text = "Password", font = ("arial", 10, "bold"), fg = "black").place(x=5,y=195)



email = StringVar()
usernameBox = Entry(root, textvariable = email, width = 25, bg = "white").place(x = 125, y = 165)
password = StringVar()
passwordBox = Entry(root, show = "*", textvariable = password, width = 25, bg = "white").place(x = 125, y = 195)

'''
Submit microsoft email/password in order to access your token that will give
you access to APIs.
'''
def submitLogin():
    
    client_id='YOUR CLIENT ID'
    authority_url = 'YOUR AUTHORITY URL'
    '''Example authority url: https://login.microsoftonline.com/WEBSITE.com '''

    
    Dataflow_APIs.connect(email.get(),password.get(),client_id,authority_url)
    
    loginResult = Label(root, text = Dataflow_APIs.guiLoginMsg, font = ("arial", 8), fg = "green").place(x=75,y=222)
    workspacesLabel4 = Label(root, text = "Workspaces", font = ("arial", 9, "bold"), fg = "black").place(x=10,y=290)
    makeGroupIdDict()

loginButton = Button(root, text = "Submit", width = 10, height = 1, bg = "lightblue", command = submitLogin).place(x = 195, y = 220)



'''
makeGroupIdDict()


Creates a global dictionary and list box for group names.

The dictionary is {name : groupId,...}

'''

def makeGroupIdDict():

    global groupIdDict,groupListBox
    groupIdDict = Dataflow_APIs.makeIdDict()
    groupNamesList = []
    
    for key in groupIdDict.keys():
        groupNamesList.append(key)





    groupFrame = Frame(root)
    groupScroll = Scrollbar(groupFrame, orient=VERTICAL)
    
    groupStr = StringVar()
    groupListBox = Listbox(groupFrame, width = 30,listvariable = groupStr, selectmode = "single",yscrollcommand=groupScroll.set)
    
    groupScroll.config(command=groupListBox.yview)
    groupScroll.pack(side=RIGHT, fill=Y)


    groupFrame.place(x=10,y=330)
    groupListBox.pack()

    
    sorted_list = sorted(groupNamesList, key=str.lower)
    sorted_list.reverse()
    
    for groupName in sorted_list:

        groupListBox.insert(0,groupName)

    


'''
selectGroup()
selectWGButton

Assigns a workspace you select to the global variable.
'''


def selectGroup():
    global workspace, openWorkspaceInBrowser
    tempWS = groupListBox.curselection()
    workspace = groupListBox.get(tempWS)
    print(workspace)
    label5 = Label(root, text = "Dataflows", font = ("arial", 10, "bold"), fg = "blue").place(x=10,y=570)
    makeDataflowIdDict()
    workspaceNote = Label(root, text = f"Current workspace: {workspace}          ", font = ("arial", 10, "bold"), fg = "red")
    workspaceNote.place(x=25,y=40)

    openWorkspaceInBrowser = Button(root, text = "Open Workspace in Browser", width = 25, height = 1, bg = "lightblue", command = open_workspace).place(x = 5, y = 5)



selectWGButton = Button(root, text = "Select Workspace", width = 15, height = 1, bg = "lightblue", command = selectGroup).place(x = 125, y = 290)


'''
Opens workspace in browser.
'''

def open_workspace():
# url format: https://app.powerbi.com/groups/e216c52e-a8dd-401a-992d-faf376110892/list/dashboards

    groupId = groupIdDict[workspace]
    url = f"https://app.powerbi.com/groups/{groupId}/list/dashboards"
    print(url)
    webbrowser.open(url)







'''
makeDataflowIdDict()
makeDataflowIdDictButton

Makes a dictionary of all the dataflows in a workspace, puts the names of the dictionaries in a list,
and uses that list to display all the dataflows in the list box.

'''
def makeDataflowIdDict():
    global dataflowIdDict,dataflowListBox
    dataflowIdDict = Dataflow_APIs.makeDataflowIdDict(groupIdDict[workspace])

    dataflowNamesList = []
    for key in dataflowIdDict.keys():
        
        dataflowNamesList.append(key)



    dataflowFrame = Frame(root)
    dataflowScroll = Scrollbar(dataflowFrame, orient=VERTICAL)
    
    dataflowStr = StringVar()
    dataflowListBox = Listbox(dataflowFrame, width = 30,listvariable = dataflowStr, selectmode = "single",yscrollcommand=dataflowScroll.set)
    

    dataflowScroll.config(command=dataflowListBox.yview)
    dataflowScroll.pack(side=RIGHT, fill=Y)

    dataflowFrame.place(x=10,y=600)
    dataflowListBox.pack()

    sorted_list = sorted(dataflowNamesList, key=str.lower)
    sorted_list.reverse()
    
    for dataflowName in sorted_list:

        dataflowListBox.insert(0,dataflowName)



'''
selectDataflow()
selectDFButton

Button to select the dataflow you're currently using.
'''
def selectDataflow():
    global dataflow
    tempDF = dataflowListBox.curselection()
    dataflow = dataflowListBox.get(tempDF)

    label2 = Label(root, text = f"Current dataflow: {dataflow}           ", font = ("arial", 10, "bold"), fg = "red")
    label2.place(x = 25, y = 60)
    print(dataflow)

selectDFButton = Button(root, text = "Select Dataflow", width = 15, height = 1, bg = "lightblue", command = selectDataflow).place(x = 120, y = 570)


'''
Refresh currently chosen dataflow.
'''

def refreshDf():

    groupId = groupIdDict[workspace]
    dataflowId = dataflowIdDict[dataflow]

    
    Dataflow_APIs.refreshDataflow(groupId,dataflowId)
    dfRefResult = Label(root, text = Dataflow_APIs.refreshMsg, font = ("arial", 7), fg = "green").place(x=325,y=150)
    
refreshDFButton = Button(root, text = "Refresh Dataflow", width = 20, height = 1, bg = "lightblue", command = refreshDf).place(x = 335, y = 175)


# ---------------------------------
# Make list box for days  -------------------------------------------------------------
# ---------------------------------
dayStr = StringVar()
dayListBox = Listbox(root, width = 30,listvariable = dayStr, selectmode = "multiple")
dayListBox.place(x=380,y=300)

listOfDays = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
listOfDays.reverse()
for day in listOfDays:
    
    dayListBox.insert(0,day)  # insert days into the list box

# ---------------------------------
# --------------------------------------------------------------------------------------
# ---------------------------------


'''
For the Select Days button to select the days in the refresh schedule.
'''
def selectDays():
    global refDays,dayStr,timeStr,refDaysNote
    timeStr = ''
    dayStr = ''
    refDays = []
    daySelection = dayListBox.curselection()
    for day in daySelection:
        refDays.append(dayListBox.get(day))  # append the selected refresh days to refDays list.

    print(refDays)


    count = 0
    for day in refDays:
        
        if count > 0:
            
            dayStr = dayStr + ', ' + day
            count = count + 1
        else:
            dayStr = dayStr + day
            count = count + 1
    dayStr = dayStr + "   "*8

    if "refDaysNote" in globals():
        
        refDaysNote.destroy()


        
        
    refDaysNote = Label(root, text = f"Selected Days: {dayStr}                       ", font = ("arial", 8, "bold"), fg = "black")
    refDaysNote.place(x=425,y=40)


selectDaysButton = Button(root, text = "Select Refresh Day(s)", width = 20, height = 1, bg = "lightblue", command = selectDays).place(x = 385, y = 470)


# ---------------------------------
# Make list box for times  -------------------------------------------------------------
# ---------------------------------
timeFrame = Frame(root)
timeScroll = Scrollbar(timeFrame, orient=VERTICAL)

timeStr1 = StringVar()
timeListBox = Listbox(timeFrame, width = 10,listvariable = timeStr1, selectmode = "multiple",yscrollcommand=timeScroll.set)


timeScroll.config(command=timeListBox.yview)
timeScroll.pack(side=RIGHT, fill=Y)


timeFrame.place(x=380,y=600)

timeListBox.pack()

listOfTimes = []

for hour in range(24):

    if hour < 10:

        strHour = str(hour)
        tempHour = f"0{strHour}:00"
        tempHourHalf = f"0{strHour}:30"
        listOfTimes.append(tempHour)
        listOfTimes.append(tempHourHalf)

    else:

        strHour = str(hour)
        tempHour = f"{strHour}:00"
        tempHourHalf = f"{strHour}:30"
        listOfTimes.append(tempHour)
        listOfTimes.append(tempHourHalf)

listOfTimes.reverse()
for time in listOfTimes:
    
    timeListBox.insert(0,time)  # insert times into the list box

# ---------------------------------
# --------------------------------------------------------------------------------------
# ---------------------------------

'''
Function for the Select Times button.
Selects the times for our update refresh schedule API and destroys/replaces the label
we used to display the time(s).
'''


def selectTimes():
    global refTimes,timeStr,label3
    timeStr = ''
    refTimes = []
    timeSelection = timeListBox.curselection()
    for time in timeSelection:
        refTimes.append(timeListBox.get(time))  # append the selected refresh times to refTimes list.

    print(refTimes)

      


    count = 0
    for t in refTimes:
        if count > 0:
            
            timeStr = timeStr + ', ' + t
            count = count + 1
        else:
            timeStr = timeStr + t
            count = count + 1

    print(timeStr)



    if "timeLabel" in globals():
        
        timeLabel.destroy()
        timeLabel.place_forget()

    if "timeLabel" in locals():
        
        timeLabel.destroy()
        timeLabel.place_forget()


    
    
    timeLabel = Label(root, text = f"Selected Times: {timeStr}                                  ", font = ("arial", 8, "bold"), fg = "black")
    timeLabel.place(x=425,y=85)

selectTimeButton = Button(root, text = "Select Refresh Time(s)", width = 20, height = 1, bg = "lightblue", command = selectTimes).place(x = 475, y = 700)

labelUTC = Label(root, text = "UTC", font = ("arial", 10, "bold"), fg = "black")
labelUTC.place(x=380,y=577)









'''

clearTime() clears the list of times for updating refresh schedule.

'''

def clearTime():

    refTimes.clear()
    if "timeLabel" in globals():
        
        timeLabel.destroy()
        timeLabel.pack_forget()
        refTimes.clear()
        timeStr = ''
        print(f"refTimes = {refTimes}")
    
    labelSpace = Label(root, text = "                             "*2, font = ("arial", 8, "bold"), fg = "black")
    labelSpace.place(x=380,y=577)



#clearButton = Button(root, text = "Clear", width = 10, height = 1, bg = "lightblue", command = clearTime).place(x = 625, y = 700)

'''
Updates the refresh schedule with the selected dates and times.
'''

def updateDataflowRefSched():

    groupId = groupIdDict[workspace]
    dataflowId = dataflowIdDict[dataflow]
    
    
    # updateRefSched(groupId,dataflowId,refDays,refTimes) where refDays is a list of days and refTimes is a list of times in UTC.
    Dataflow_APIs.updateRefSched(groupId,dataflowId,refDays,refTimes)

    refSchedResult = Label(root, text = Dataflow_APIs.schedMsg, font = ("arial", 7), fg = "green").place(x=530,y=150)




updateRefSchedButton = Button(root, text = "Update Refresh Schedule", width = 20, height = 1, bg = "lightblue", command = updateDataflowRefSched).place(x = 535, y = 175)






root.mainloop()
