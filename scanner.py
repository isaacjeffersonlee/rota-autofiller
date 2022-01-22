# from microsoftgraph.client import Client
from oauthlib.oauth2 import WebApplicationClient
from urllib.parse import urlparse, parse_qs
import requests
import json
import msal
import subprocess
import webbrowser
import time
from datetime import datetime as dt
import pyautogui as pa

class GraphClient:
    AUTHORITY_URL = 'https://login.microsoftonline.com/'
    BASE_URL = 'https://graph.microsoft.com/v1.0'
    AUTH_ENDPOINT = '/oauth2/v2.0/authorize?'
    TOKEN_ENDPOINT = '/oauth2/v2.0/token'
    STATE = "1234" # Can be any string

    def __init__(
            self,
            client_id: str,
            client_secret: str,
            redirect_uri: str,
            scope: list[str],
            account_type: str,
            root_driveid: str,
            ):
        """Initialize the Graph API client."""
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.scope = scope
        self.account_type = account_type
        self.root_driveid = root_driveid
        self.access_token = None
        self.refresh_token = None
        self.delta_token = None
        # Initialize the ConfidentialClientApplication object
        self.client_app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=self.AUTHORITY_URL + self.account_type,
            client_credential=self.client_secret,
        )

    def get_access_code(self):
        """Get web application client flow access code."""
        auth_url = self.client_app.get_authorization_request_url(
            scopes=self.scope, state=self.STATE, redirect_uri=self.redirect_uri
            )
        print("-" * 56) # Aesthetics
        print("Paste this into your browser and copy the resulting url: ")
        print("")
        print(auth_url)
        print("")
        unparsed_url = input("Url: ")
        parsed_url = urlparse(unparsed_url)
        print("-" * 56)
        return parse_qs(parsed_url.query)['code'][0]

    def get_access_token(self):
        """Get the access token and refresh token."""
        code = self.get_access_code()
        token_dict = self.client_app.acquire_token_by_authorization_code(
            code=code, scopes=self.scope, redirect_uri=self.redirect_uri
        )
        self.access_token = token_dict['access_token']
        self.refresh_token = token_dict['refresh_token']

    def refresh_access_token(self):
        """Get a new access token using our refresh token."""
        # Grab a new token using our refresh token.
        token_dict = self.client_app.acquire_token_by_refresh_token(
            refresh_token=self.refresh_token, scopes=self.scope
        )
        self.access_token = token_dict['access_token']
        self.refresh_token = token_dict['refresh_token']

    def get_rota(self, relative_file_path):
        """
        Get the json response from requesting the rota at relative_file_path.
        (Relative to the GraphClient objects root_driveid).
        """
        headers = { "Authorization": f"Bearer {self.access_token}" }
        request_url = self.BASE_URL + f"/drives/{self.root_driveid}/root:{relative_file_path}"
        r = requests.get(url=request_url, headers=headers)
        return r.json()

    def get_delta_token(self):
        """
        Get the delta token for the objects root drive.
        This is a string which is used in the request url to 
        get any new changes in the drive since the last delta
        token was issued.
        """
        headers = { "Authorization": f"Bearer {self.access_token}" }
        request_url = self.BASE_URL + f"/drives/{self.root_driveid}/root/delta"
        r = requests.get(url=request_url, headers=headers).json()
        delta_link = None
        while not delta_link: # Keep going until we find the last page
            try:
                delta_link = r['@odata.deltaLink']
            except KeyError: # More pages present
                next_link = r['@odata.nextLink']
                next_token = parse_qs(urlparse(next_link).query)['token'][0]
                request_url = self.BASE_URL + f"/drives/{self.root_driveid}/root/delta(token={next_token})"
                r = requests.get(url=request_url, headers=headers).json()

        delta_token = parse_qs(urlparse(delta_link).query)['token'][0]
        self.delta_token = delta_token

    def check_for_new(self):
        """Update delta_token, using old delta_token."""
        if not self.delta_token: # If we don't yet have a delta token
            self.get_delta_token()
        headers = { "Authorization": f"Bearer {self.access_token}" }
        request_url = self.BASE_URL + f"/drives/{self.root_driveid}/root/delta(token='{self.delta_token}')"
        r = requests.get(url=request_url, headers=headers)
        return r.json()



def scroll_to_top_left(scroll_length=30):
    """Using the arrow keys, scroll to the top left of the screen."""
    pa.press('left', presses=scroll_length)
    pa.press('pageup')


def scan_for_new_rota(root_drive='personal'):
    with open('credentials.json', 'r') as f: # Read in our credentials json
        credentials = json.load(f)
    scope=['User.Read', 'Files.ReadWrite.All', 'Files.Read.All',
            'Sites.Read.All', 'Sites.Manage.All', 'Sites.ReadWrite.All']
    redirect_uri = credentials['redirect_uri']
    client_secret = credentials['client_secret']
    client_id = credentials['client_id']
    account_type = credentials['account_type'] 
    # Beit-Bars Rotas site drive id, found by using requests to the Beit-Bars site
    rota_driveid = 'b!eeW687o3gEGiBnBABZArSmR-zAXXs-xPldmu_FiTD2UJoNtojn0aR7IRk7293MHO'
    # My personal drive id
    personal_driveid = 'b!VVp9GUD9v06cyfM41FM4_CwrynnMOI5LpvhML0mbJnIpc8sEVKhASL9LnZIC63dh'
    # Instantiate a GraphClient object
    if root_drive == 'personal':
        root_driveid = personal_driveid
    elif root_drive == 'rota':
        root_drive = rota_driveid
    else:
        print(f"Error, {root_drive} is not a valid drive!")

    gc = GraphClient(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            scope=scope,
            account_type=account_type,
            root_driveid=root_driveid) # <<<<<<<<<<<<<<<<<<

    gc.get_access_token() # Web app client authentication

    counter = 0
    start_time = dt.now() # Start duration timer
    with open('old_rotas.txt', 'r') as f: # Read in old rota names as a list
        old_rotas = f.read().splitlines()

    switch = True
    while switch:
        time.sleep(5) # Zzzzz
        gc.refresh_access_token() # Get new access token, using refresh token
        r = gc.check_for_new()
        # try:
        #     if r['value']: print(f"Change detected in {r['value'][-1]['name']}")
        # except KeyError:
        #     print('KeyError detected!')
        time_elapsed = str(dt.now() - start_time).split('.')[0]
        bar_length = 40
        arrow_length = counter % int(bar_length / 2)
        print(f"   Scanning {root_drive} drive for new rotas...  ")
        print("")
        print(f"    Check number: {counter}")
        print(f"    Time elapsed: {time_elapsed}")
        print("")


        for change in r['value']:
            try:
                # Filter only excel files 
                if change['file']['mimeType'] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                    name = change['name']
                    year = name.split(".")[0][-2:] # Final two characters
                    # We are only interested in files for 2022
                    if year == '22' and name not in old_rotas:
                        url = change['webUrl']
                        print(f"New excel file detected: {name}")
                        print(f"Created: {change['createdDateTime']}")
                        print(f"Last modified: {change['lastModifiedDateTime']}")
                        print("Opening file in browser!")
                        webbrowser.open(url, new=0, autoraise=True) # open in browser
                        time.sleep(3)
                        scroll_to_top_left(scroll_length=30)
                        print("Scrolling to top left")

                        switch = False # exit while loop

            except KeyError:
                pass

        counter += 1 # Increment counter
        print("-" * 56) # Aesthetics

    print(f"Finished scanning, found {name}")

    return name


if __name__ == '__main__':
    scan_for_new_rota()
