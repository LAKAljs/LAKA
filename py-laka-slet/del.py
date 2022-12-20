#-*- coding: utf-8 -*-
# Define imports
import json
import msal
import requests

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
headers = {
  'Authorization': access_token
}

# Make a GET request to the provided url, passing the access token in a header

data = json.load(open('rescal.json'))

for elem in data:
    url = 'https://graph.microsoft.com/v1.0/users/alm@laka.dk/events/' + elem['id']
    data = requests.delete(url=url, headers=headers)


