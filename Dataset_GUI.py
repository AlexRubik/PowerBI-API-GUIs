'''
GUI for PowerBI Dataset APIs (not all of them)
Contact: dewittat@g.cofc.edu
'''

from tkinter import *
from tkinter import ttk
import Dataflow_APIs
import Dataset_APIs
import webbrowser

root = Tk()
root.title("PowerBI Datasets GUI")
root.geometry("870x790+450+100")
heading = Label(root, text = "PowerBI Datasets GUI", font = ("arial",25,"bold"),fg = "green").pack()

 

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
    Dataset_APIs.resultFunction(Dataflow_APIs.result)  # Telling Dataset_APIs the value of the result variable.
    loginResult = Label(root, text = Dataflow_APIs.guiLoginMsg, font = ("arial", 8), fg = "green").place(x=75,y=222)
    label4 = Label(root, text = "Workspaces", font = ("arial", 9, "bold"), fg = "black").place(x=10,y=290)
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
    global workspace, openWorkspaceInBrowser, groupId

    
    tempWS = groupListBox.curselection()
    workspace = groupListBox.get(tempWS)

    groupId = groupIdDict[workspace]
    
    print(workspace)
    datasetsLabel5 = Label(root, text = "Datasets", font = ("arial", 10, "bold"), fg = "blue").place(x=10,y=570)
    makeDatasetIdDict()
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
Makes a dictionary of all the datasets in a workspace, puts the names of the dictionaries in a list,
and uses that list to display all the datasets in the list box.

'''
def makeDatasetIdDict():
    global datasetIdDict,datasetListBox
    datasetIdDict = Dataset_APIs.getDatasetsInGroup(groupIdDict[workspace])

    datasetNamesList = []
    for key in datasetIdDict.keys():
        
        datasetNamesList.append(key)


    datasetFrame = Frame(root)
    datasetScroll = Scrollbar(datasetFrame, orient=VERTICAL)
    
    datasetStr = StringVar()
    datasetListBox = Listbox(datasetFrame, width = 30,listvariable = datasetStr, selectmode = "single",yscrollcommand=datasetScroll.set)
    

    datasetScroll.config(command=datasetListBox.yview)
    datasetScroll.pack(side=RIGHT, fill=Y)

    datasetFrame.place(x=10,y=600)
    datasetListBox.pack()



#-----------------------------------------------------------------



    sorted_list = sorted(datasetNamesList, key=str.lower)
    sorted_list.reverse()
    for datasetName in sorted_list:

        datasetListBox.insert(0,datasetName)



'''
Button to select the dataset you're currently using.
'''
def selectDataset():
    global dataset, datasetId
    tempDS = datasetListBox.curselection()
    dataset = datasetListBox.get(tempDS)
    datasetId = datasetIdDict[dataset]

    label2 = Label(root, text = f"Current dataset: {dataset}           ", font = ("arial", 10, "bold"), fg = "red")
    label2.place(x = 25, y = 60)
    print(dataset)

selectDSButton = Button(root, text = "Select Dataset", width = 15, height = 1, bg = "lightblue", command = selectDataset).place(x = 120, y = 570)




'''
Refresh currently chosen dataset.
'''

def refreshDs():

    groupId = groupIdDict[workspace]
    datasetId = datasetIdDict[dataset]

    
    Dataset_APIs.refreshDataset2(groupId,datasetId)
    dsRefResult = Label(root, text = Dataset_APIs.refreshMsg, font = ("arial", 7), fg = "green").place(x=325,y=150)
    
refreshDsButton = Button(root, text = "Refresh Dataset", width = 20, height = 1, bg = "lightblue", command = refreshDs).place(x = 335, y = 175)








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
# Make list box for times  ---------------------------------------------------------------------------------------------------
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
# -------------------------------------------------------------------------------------------------------------
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
def updateDatasetRefSched():

    groupId = groupIdDict[workspace]
    datasetId = datasetIdDict[dataset]
    
    
    # updateRefSched(groupId,datasetId,refDays,refTimes) where refDays is a list of days and refTimes is a list of times in UTC.
    Dataset_APIs.updateRefSched(groupId,datasetId,refDays,refTimes)

    refSchedResult = Label(root, text = Dataset_APIs.schedMsg, font = ("arial", 7), fg = "green").place(x=530,y=150)




updateRefSchedButton = Button(root, text = "Update Refresh Schedule", width = 20, height = 1, bg = "lightblue", command = updateDatasetRefSched).place(x = 535, y = 175)


'''
REFRESH HISTORY AND SCHEDULE

An extra window to view the current schedule and refresh history of
the dataset.
'''
def refHistoryWindow():
    global numberRecords,top

    top = Toplevel()
    top.geometry("890x890+470+60")
    numberRecords = StringVar()
    numberBox = Entry(top, textvariable = numberRecords, width = 12, bg = "white").place(x = 250, y = 15)
    
    labelQ = Label(top, text = "How many records would you like to see?", font = ("arial", 7, "bold"), fg = "black")
    labelQ.place(x=5,y=15)


    schedStr = Dataset_APIs.getRefSched(groupId,datasetId)


    
    labelSched = Label(top, text = schedStr, font = ("arial", 8, "bold"), fg = "black")
    labelSched.place(x=5,y=45)

    labelDataset = Label(top, text = f"Dataset: {dataset}    ", font = ("arial", 12, "bold"), fg = "red")
    labelDataset.place(x = 640, y = 15)
    
    getRefHistoryButton = Button(top, text = "Show History", width = 23, height = 1, bg = "yellow", command = getRefHis).place(x = 360, y = 15)
    


refHistoryButton = Button(root, text = "Refresh History/Schedule", width = 25, height = 1, bg = "yellow", command = refHistoryWindow).place(x = 535, y = 235)

'''
Shows refresh history in text in the new window.

'''


def getRefHis():

    global t,v




    try:
        
        if t.winfo_exists() == 1:

            t.destroy()
            v.destroy()

    except:
        print("t widget hasn't been created yet")

        
        
    


    v = Scrollbar(top)
    v.pack(side = RIGHT, fill = Y)

    t = Text(top, width = 15, height = 45, wrap = NONE, yscrollcommand = v.set)

    numRecords = numberRecords.get()
    historyStr = Dataset_APIs.getRefHistory(groupId,datasetId,numRecords)
    t.insert(END,historyStr)


    t.pack(side=BOTTOM, fill=X)
    v.config(command=t.yview)



    
'''
A small window that shows ONLY the refresh schedule of the selected dataset.

'''



def getRefSchedWindow():
    
    schedWindow = Toplevel()
    schedWindow.geometry("500x175+470+160")

    schedStr = Dataset_APIs.getRefSched(groupId,datasetId)


    
    labelSched = Label(schedWindow, text = schedStr, font = ("arial", 8, "bold"), fg = "black")
    labelSched.place(x=5,y=15)

#refSchedButton = Button(root, text = "Current Refresh Schedule", width = 25, height = 1, bg = "yellow", command = getRefSchedWindow).place(x = 535, y = 265)




root.mainloop()


