'''
Functions to pull/send data to dataflow APIs.

Contact: dewittat@g.cofc.edu

Thanks to Ben Watt at datalineo.com for the Microsoft Authentication code.

https://www.datalineo.com/post/power-bi-rest-api-with-python-and-microsoft-authentication-library-msal

'''



import msal
import requests
import json
import Dataflow_APIs
import time

'''
resultFunction

I made this function so that I could assign "result" a value from another program.
The result variable is needed and used by almost every function and only needs
to be pulled once. When referring to Dataflow_APIs in Dataset_GUI, I call a
function from Dataflow_APIs that pulls the result (token) using msal. In order to
inititialize "result" in Dataset_APIs, I needed to make this function so it can be
used in this program. As opposed to calling for another token within Dataset_APIs.


'''
def resultFunction(result1):

    global result
    result = result1
    

def connect(email,pword,client_id,authority_url):
    global app,result,guiLoginMsg
    
    
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
makeDatasetDict()

Returns a dictionary of datasets from "My Workspace".  name : id

https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/getdatasets

'''
def makeDatasetDict():  

    global datasetDict
    datasetDict = dict()
    url = 'https://api.powerbi.com/v1.0/myorg/datasets'
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        api_out = requests.get(url=url, headers=header)
        

        
    else:
        print(result.get("error"))
        print(result.get("error_description"))



    initDict = api_out.json()
    initDict = initDict['value']
    

    for dic in initDict:   # iterating through the list of dictionaries
        
        ID = dic['id']
        name = dic['name']
        datasetDict[name] = ID

    return datasetDict





'''
refreshDataset(datasetId)

Triggers a refresh for the specified dataset from "My Workspace".

In Shared capacities this call is limited to eight times per day (including refreshes executed via Scheduled Refresh)
In Premium capacities this call is not limited in number of times per day, but only by the available resources in the capacity,
hence if overloaded, the refresh execution may be throttled until the load is reduced. If this throttling exceeds 1 hour, the refresh will fail.

POST https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/refreshes

https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/refreshdataset

# NOTE: This function will only work for a dataset in My Workspace !
# refreshDataset2() let's you choose the workspace (groupId) that you need to refresh from.

'''

testDataset = 'Item Maintenance Process Report beta'

def refreshDataset(datasetId):
    

    url = f"https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/refreshes"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        data = {"refreshRequest":"y"}
        api_out = requests.post(url=url, json=data, headers=header)
        
        if api_out.status_code == 202:
            print(f"Successful refresh of dataset: {testDataset} <> {datasetId}")
        return api_out.status_code
    else:
        print(result.get("error"))
        print(result.get("error_description"))

    






'''
refreshDataset2(groupId, datasetId)

Triggers a refresh for the specified dataset (datasetId) from the specified workspace (groupId).

https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/refreshdatasetingroup

# POST https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes
'''

testDataset2 = 'xx'
def refreshDataset2(groupId, datasetId):
    global refreshMsg
    
    t = time.localtime()
    current_time = time.strftime("%H:%M:%S", t)
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        data = {"refreshRequest":"y"}
        api_out = requests.post(url=url, json=data, headers=header)
        
        if api_out.status_code == 202:
            print(f"Successful refresh of dataset: {datasetId}")
            refreshMsg = f"Successful Refresh @ {current_time} local time"
        else:
            code = str(api_out.status_code)
            refreshMsg = f"Failed to refresh. Status code {code}"
            print(api_out.status_code)
        return api_out.status_code
    else:
        refreshMsg = "No access token in result"
        print(result.get("error"))
        print(result.get("error_description"))


'''
getDatasetsInGroup(groupId)

Returns a dictionary of datasets from the specified groupId in the format name : id

# GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets
https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/getdatasetsingroup
'''
def getDatasetsInGroup(groupId):
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        
        api_out = requests.get(url=url, headers=header)
        
        if api_out.status_code == 200:
            print("Success")
            jsonDict = api_out.json()
            jsonDict = jsonDict['value']
            


    else:
        print(api_out.status_code)
        print(result.get("error"))
        print(result.get("error_description"))


    datasetDict1 = dict()
    for dic in jsonDict:   # iterating through the list of dictionaries and making a new one with name : id
        
        ID = dic['id']
        name = dic['name']
        datasetDict1[name] = ID
        
    
    
    return datasetDict1




'''




https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/updaterefreshscheduleingroup
PATCH https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule
'''



def updateRefSched(groupId,datasetId,refDays,refTimes): 

    global schedMsg

    t = time.localtime()
    current_time = time.strftime("%H:%M:%S", t)
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        data = {
  "value": {
    "days": refDays,
    "times": refTimes,
    "localTimeZoneId": "UTC",
    
  }
}
        api_out = requests.patch(url=url, json=data, headers=header)
        
        if api_out.status_code == 200:
            schedMsg = f"Refresh Schedule Updated Successfully @ {current_time} local time"
            print(f"Successful refresh schedule update of dataset:  <> {datasetId}")

        else:
            code = str(api_out.status_code)
            schedMsg = f"Failed. Code {code}                                               "

        return api_out.status_code
    else:
        schedMsg = "No token in result"
        print(result.get("error"))
        print(result.get("error_description"))





'''
Returns the refresh history of the specified dataset from the specified workspace.


https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/getrefreshhistoryingroup
GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes

or

GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes?$top={$top}
'''


def getRefHistory(groupId,datasetId,numRecords):
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes?$top={numRecords}"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        
        api_out = requests.get(url=url, headers=header)
        
        if api_out.status_code == 200:
            print("Success")
            jsonDict = api_out.json()
            print()
            print("--"*100)
            
            jsonDict = jsonDict['value']
            
            


    else:
        print(api_out.status_code)
        print(result.get("error"))
        print(result.get("error_description"))


    refHisStr = ''
    for dic in jsonDict:   # iterating through the list of dictionaries and making a new one with name : id
        refHisStr = refHisStr + "==========="*10 + '\n' + "==========="*10 + '\n'
        for key in dic:

            tempKeyValPair = str(key) + ' = ' + str(dic[key]) + '\n'
            refHisStr = refHisStr + tempKeyValPair
        


    

    
    return refHisStr



'''

https://docs.microsoft.com/en-us/rest/api/power-bi/datasets/getrefreshscheduleingroup
GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule
'''

def getRefSched(groupId,datasetId):

    url = f"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule"
    
    if 'access_token' in result:
        access_token = result['access_token']
        header = {'Content-Type':'application/json','Authorization': f'Bearer {access_token}'}
        
        api_out = requests.get(url=url, headers=header)
        
        if api_out.status_code == 200:
            print("Success")
            jsonDict = api_out.json()
            print()
            
            del jsonDict['@odata.context']
            print(jsonDict)
            
            


    else:
        print(api_out.status_code)
        print(result.get("error"))
        print(result.get("error_description"))


    refSchedStr = ''
    for key in jsonDict:
        
        tempKeyValPair = str(key) + ' = ' + str(jsonDict[key]) + '\n'
        refSchedStr = refSchedStr + tempKeyValPair


    return refSchedStr






def main():
    
    '''
    Test functions...
    '''

    

#main()

    












    
