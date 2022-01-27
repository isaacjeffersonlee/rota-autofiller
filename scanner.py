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
    """
    Microsoft Graph API client class.
    
    Attributes
    ----------
    client_id : str
        API app client id.

    client_secret : str
        API app client secret.

    redirect_uri : str
        Redirect uri/url for authentication.

    scope : list[str]
        Permissions required.

    account_type : str
        E.g 'organizations' or 'personal'.

    root_driveid : str
        The unique microsoft graph driveItem id of the root
        directory to search.
    """
    
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
        self.token_expires_in = None
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
        webbrowser.open(auth_url, new=0, autoraise=True)
        unparsed_url = input("Url: ")
        parsed_url = urlparse(unparsed_url)
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
        self.token_expires_in = token_dict['expires_in'] # In seconds

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

    def get_driveItems(self, path_relative_to_root):
        """Get a list of the driveItems at the relative_path to the root drive."""
        headers = { "Authorization": f"Bearer {self.access_token}" }
        request_url = self.BASE_URL + f"/drives/{self.root_driveid}/root:/{path_relative_to_root}:/children"
        r = requests.get(url=request_url, headers=headers)
        return r.json()

