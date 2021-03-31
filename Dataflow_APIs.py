'''
Functions to pull/send data to dataflow APIs

Contact: dewittat@g.cofc.edu

Thanks to Ben Watt at datalineo.com for the Microsoft Authentication code.

https://www.datalineo.com/post/power-bi-rest-api-with-python-and-microsoft-authentication-library-msal

'''


import msal
import requests
import json
import time


def connect(email,pword,client_id,authority_url):   # returns result
    global app,result,guiLoginMsg
    # --------------------------------------------------
    # Set local variables
    # --------------------------------------------------
    
    username = str(email)
    password = str(pword)
     
    
    scope = ["https://analysis.windows.net/powerbi/api/.default"]
    url_groups = 'https://api.powerbi.com/v1.0/myorg/groups'

    # --------------------------------------------------
    # Use MSAL to grab a token
    # --------------------------------------------------
    app = msal.PublicClientApplication(client_id, authority=authority_url)
    result = app.acquire_token_by_username_password(username=username,password=password,scopes=scope)

    if 'access_token' in result:
        print("Successful login")
        guiLoginMsg = "Successful login       "
        

    else:
        
        guiLoginMsg = "Wrong login credentials"
        print(result.get("error"))
        print(result.get("error_description"))

    return result

    





'''
makeIdDict()

Returns a dictionary of workspace names and their ids  (groupIds)


https://docs.microsoft.com/en-us/rest/api/power-bi/groups/getgroups

GET https://api.powerbi.com/v1.0/myorg/groups


'''
 
def makeIdDict():   # for groups/workspaces
    url_groups = 'https://api.powerbi.com/v1.0/myorg/groups'
    global groupsDict   # dictionary of name : id   for the groups
    groupsDict = dict()
    if 'access_token' in result:
        
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        api_out = requests.get(url=url_groups, headers=header)

        jsonGroupsDict = api_out.json()
        listOfGroupDicts = jsonGroupsDict['value']
    else:
        print(result.get("error"))
        print(result.get("error_description"))

    
    for dic in listOfGroupDicts:   # iterating through the list of dictionaries
        
        ID = dic['id']
        name = dic['name']
        groupsDict[name] = ID

    return groupsDict
            







'''

makeDataflowIdDict(groupId)

Returns a list of all dataflows from the specified workspace.

https://docs.microsoft.com/en-us/rest/api/power-bi/dataflows/getdataflows
GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/dataflows
'''



def makeDataflowIdDict(groupId):
    global dataflowDict
    dataflowDict = dict()     # dictionary of Dataflow name : objectId
    
    url = 'https://api.powerbi.com/v1.0/myorg/groups/' + groupId + '/dataflows'
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        api_out = requests.get(url=url, headers=header)
        dataflowJson = api_out.json()
        listOfDataflowDicts = dataflowJson['value']
        if api_out.status_code == 200:
            print("Successful get request.")
    else:
        print(result.get("error"))
        print(result.get("error_description"))
    

    for dic in listOfDataflowDicts:   # iterating through the list of dictionaries
        
        ID = dic['objectId']
        name = dic['name']
        dataflowDict[name] = ID

    return dataflowDict




'''
refreshDataflow(groupId, dataflowId)

Triggers a refresh for the specified dataflow from the specified workspace.

https://docs.microsoft.com/en-us/rest/api/power-bi/dataflows/refreshdataflow
POST https://api.powerbi.com/v1.0/myorg/groups/{groupId}/dataflows/{dataflowId}/refreshes?processType={processType}
'''

testDataflow = 'BOM'

def refreshDataflow(groupId, dataflowId):  # parameters are strings

    global refreshMsg
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/dataflows/{dataflowId}/refreshes?processType=default"
    t = time.localtime()
    current_time = time.strftime("%H:%M:%S", t)
            
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        data = {"refreshRequest":"y"}
        api_out = requests.post(url=url, json=data, headers=header)

        if api_out.status_code == 200:
            
            print(f"Successful refresh of dataflow:  <> {dataflowId}")
            refreshMsg = f"Successful Refresh @ {current_time} local time"

        else:
            code = str(api_out.status_code)
            refreshMsg = f"Failed to refresh. Status code {code}"
        return api_out.status_code
    else:
        refreshMsg = "No access token in result"
        print(result.get("error"))
        print(result.get("error_description"))





'''
updateRefSched(groupId,dataflowId)

Creates or updates the specified dataflow refresh schedule configuration.

Refer to the format of the json data in the docs.

https://docs.microsoft.com/en-us/rest/api/power-bi/dataflows/updaterefreshschedule
PATCH https://api.powerbi.com/v1.0/myorg/groups/{groupId}/dataflows/{dataflowId}/refreshSchedule
NOTE: Don't change the timezone. I tried to change it to EST and it didn't support it.

days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
refDays = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
refTimes = ["10:00", "16:00"]
'''


def updateRefSched(groupId,dataflowId,refDays,refTimes): 

    global schedMsg

    t = time.localtime()
    current_time = time.strftime("%H:%M:%S", t)
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/dataflows/{dataflowId}/refreshSchedule"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        data = {
  "value": {
    "days": refDays,
    "times": refTimes,
    "enabled": True,
    "localTimeZoneId": "UTC",
    "notifyOption": "NoNotification"
  }
}
        api_out = requests.patch(url=url, json=data, headers=header)
        
        if api_out.status_code == 200:
            schedMsg = f"Refresh Schedule Updated Successfully @ {current_time} local time"
            print(f"Successful refresh schedule update of dataflow:  <> {dataflowId}")

        else:
            code = str(api_out.status_code)
            schedMsg = f"Failed. Code {code}                                               "

        return api_out.status_code
    else:
        schedMsg = "No token in result"
        print(result.get("error"))
        print(result.get("error_description"))









def main():
    '''
    Test functions...
    '''
    



#main()












        

