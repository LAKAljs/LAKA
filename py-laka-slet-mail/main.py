#-*- coding: utf-8 -*-
# Define imports
import json
import msal
import requests
import configparser
import re
valueArr = []

# Enter the details of your AAD app registration
client_id = '54c69a91-e394-472d-95c4-e40e38f7d7f9'
client_secret = 'ZeY8Q~8aRRKOefZGvJ0JTv215kRj50OrVcMzcbWT'
authority = 'https://login.microsoftonline.com/f2e253d3-223d-4221-9f4e-26a8f31c7bd6'
scope = ['https://graph.microsoft.com/.default']

# Create an MSAL instance providing the client_id, authority and client_credential parameters
client = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)

# First, try to lookup an access token in cache
token_result = client.acquire_token_silent_with_error(scope, account=None)

# If the token is available in cache, save it to a variable
if token_result:
  access_token = 'Bearer ' + token_result['access_token']
  print('Access token was loaded from cache')

# If the token is not available in cache, acquire a new one from Azure AD and save it to a variable
if not token_result:
  token_result = client.acquire_token_for_client(scopes=scope)
  access_token = 'Bearer ' + token_result['access_token']
  print('New access token was acquired from Azure AD')

# Copy access_token and specify the MS Graph API endpoint you want to call.
url = 'https://graph.microsoft.com/v1.0/users'
headers = {
  'Authorization': access_token
}

# Read from config-file "target.cfg" to obtain what user and period we're targetting.
config = configparser.RawConfigParser()
config.read('target.cfg')

name = dict(config.items('name'))

period = dict(config.items('period'))

print(name['name'])

# Make a GET request to the provided url, passing the access token in a header
graph_result = requests.get(url=url, headers=headers)


def getDN(graph_result):
  # Decodes from python obj to json string in order to get correct ascii representation.
  graph_decoded = json.dumps(graph_result.json(), ensure_ascii=False)

  # And then back to python object targeting the "value" field in the object.
  vals = json.loads(graph_decoded)["value"]

  # Try to get the next page of results.
  try:
    odataNext = json.loads(graph_decoded)['@odata.nextLink']
  except:
    odataNext = False

  # iterates through the values, and appends the result to the end of the ValueArr array object.
  for obj in vals:
    valueArr.append({'DP': obj["displayName"], 'mail': obj["id"]})

  # if there was another page of results, obtain it and run this funtion again.
  if odataNext: 
    getDN(requests.get(url=odataNext, headers=headers))

  #Open user.json, in order to write to it.
  with open("users.json", "w", encoding='utf-8') as outfile:
      # writes the value of valueArr to it.
      json.dump(valueArr, outfile, ensure_ascii=False, separators=(', \n', ":"))

  # Run next methos with the name from "target.cfg"
  getCal(name['name'])

  # Then exit
  exit()

def use_regex(input_text):
    # Checks for a pattern that matches with the name and Danish SSN.
    pattern = re.compile(r"([0-9]{10})", re.IGNORECASE)
    if pattern.search(input_text) != None:
      return True
    else:
      return False


def getCal(peeps):
  # Defies what url to call.
  calUrl = 'https://graph.microsoft.com/v1.0/users/' + peeps + '/messages?startdatetime=' + period['start'] + '&enddatetime=' + period['end'] + '&$top=400000'
  odataNext = True
  values = []

  #If there is a next page
  while odataNext:
    #get request for the call url
    graph_res = requests.get(url=calUrl, headers=headers)

    #dumps to ensure ascii representation.
    graph_de = json.dumps(graph_res.json(), ensure_ascii=False)

    # converts back to python object.
    vals_de = json.loads(graph_de)


    # for elem in vals_de:
    #   #print(elem['bodyPreview'])
    #   #For each element in the python object, if the pattern matches we append it the end of values array.
    #   if use_regex(elem['bodyPreview']):
    #     values.append({'id': elem['id'], 'sub': elem['bodyPreview'].replace('\r\n', ' ')})
    
    # try:
    #   odataNext = json.loads(graph_de)['@odata.nextLink']
    #   calUrl = odataNext
    # except:
    #   odataNext = False
  
  with open("rescal.json", "w", encoding='utf-8') as outfile:
      json.dump(vals_de, outfile, ensure_ascii=False, separators=(', \n', ":"))  

#In order to start process.
getDN(graph_result)