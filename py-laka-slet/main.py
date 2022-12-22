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

# Copy access_token and specify the MS Graph API endpoint you want to call, e.g. 'https://graph.microsoft.com/v1.0/groups' to get all groups in your organization
url = 'https://graph.microsoft.com/v1.0/users'
headers = {
  'Authorization': access_token
}

config = configparser.RawConfigParser()
config.read('target.cfg')

name = dict(config.items('name'))

print(name['name'])

# Make a GET request to the provided url, passing the access token in a header
graph_result = requests.get(url=url, headers=headers)


def getDN(graph_result):
  graph_decoded = json.dumps(graph_result.json(), ensure_ascii=False)

  vals = json.loads(graph_decoded)["value"]

  try:
    odataNext = json.loads(graph_decoded)['@odata.nextLink']
  except:
    odataNext = False

  for obj in vals:
    valueArr.append({'DP': obj["displayName"], 'mail': obj["id"]})

  if odataNext: 
    getDN(requests.get(url=odataNext, headers=headers))

  with open("res.json", "w", encoding='utf-8') as outfile:
      json.dump(valueArr, outfile, ensure_ascii=False, separators=(', \n', ":"))
  
  getCal(name['name'])
  exit()

def use_regex(input_text):
    pattern = re.compile(r"[a-z]+/([0-9]{10})", re.IGNORECASE)
    if pattern.search(input_text) != None:
      return True
    else:
      return False


def getCal(peeps):
  calUrl = 'https://graph.microsoft.com/v1.0/users/' + peeps + '/calendarview?startdatetime=2022-10-18&enddatetime=2022-12-20&$top=400000'#$search="from:s_CRMprod_booking"&
  odataNext = True
  values = []

  while odataNext:
    graph_res = requests.get(url=calUrl, headers=headers)

    graph_de = json.dumps(graph_res.json(), ensure_ascii=False)
    vals_de = json.loads(graph_de)["value"]

    for elem in vals_de:
      print(elem['subject'] + ", " + elem['start']['dateTime'])

      if use_regex(elem['bodyPreview']):
        values.append({'id': elem['id'], 'sub': elem['bodyPreview'], 'from': elem['organizer']['emailAddress']['name'], 'date': elem['start']['dateTime']})
    
    try:
      odataNext = json.loads(graph_de)['@odata.nextLink']
      calUrl = odataNext
    except:
      odataNext = False

  with open("rescal.json", "w", encoding='utf-8') as outfile:
      json.dump(values, outfile, ensure_ascii=False, separators=(', \n', ":"))  

getDN(graph_result)





