#import libraries
import requests
import json
import pandas as pd
import sys
#Change path to location of Admin File - only required if you require to add a specific path for interpreter to search - otherwise comment line below
sys.path.append('C:\\yourdive')
import Admin 

class Graph_API(): 
    """A class to query the Microsoft Graph API - Created by Tiago Valente - https://www.linkedin.com/in/tiago-valente-49424968/
    
    Step 1. In Azure Portal - under Active Directory - create an app with required permissions
    Step 2. Retain client_secret, client_id, Tenant ID store securely 
    Step 3. Within App - under Authentication - Create a Redirect URI - 'https://login.microsoftonline.com/{Tenant ID}/oauth2/token'
    Step 4. In Admin file - paste your client_secret, client_id, redirect URI (token_url) - store this securely
    Step 5. Run Class Object as below
    """

    def __init__(self, token_url, grant_type, client_id,client_secret,resource,scope):
        """Initialize attributes"""
        self.token_url = token_url
        self.grant_type = grant_type
        self.client_id = client_id
        self.client_secret = client_secret
        self.resource = resource
        self.scope = scope
        
    def Get_Graph_Token(self):
        """ This function will return the token required to authenticate calls made with the Microsoft Graph API"""
        
        #Dictionary object required for Post Request
        token_data = {
         'grant_type': self.grant_type,
         'client_id': self.client_id,
         'client_secret': self.client_secret,
         'resource': self.resource,
         'scope':self.scope,
         }
        
        token_r = requests.post(self.token_url, data=token_data) #Post request serving URL and token_data
        token = token_r.json().get('access_token') #Get Request to retrieve token after post request
        
        return token
    
    def Endpoint_Call(endpoint,token):
        """ This function will return an API Status Code and API Response from the call being made to the Microsoft Graph API"""
        
        #Dictionary object required for headers in get request
        headers = {
          'Content-Type':
        'application/json',
          'Authorization': 'Bearer ' + token
        }
        
        resp = requests.get(endpoint, headers=headers)
        
        response_data = json.loads(resp.text)
        
        return resp.status_code,response_data
        
#Create object instance of Graph API using the required credentials
ClassObject = Graph_API(Admin.token_url,'client_credentials',Admin.client_id, Admin.client_secret, 'https://graph.microsoft.com', 'https://graph.microsoft.com')

#Return token values from object instance - only required if you want to see your token
#token = Graph_API.Get_Graph_Token(ClassObject)

#Retrieve JSON Response from Endpoint call 
Response = Graph_API.Endpoint_Call("https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')",Graph_API.Get_Graph_Token(ClassObject))