# Rota Bot

#### Date: 15/01/2022

### Problem
I work for my Universities student bar and we have over 100
employees. Each Monday a new rota is released on Microsoft Sharepoint
and you fill your name next to the shift you want.
With the large number of employees, the shifts get taken extremely
quickly so it is often hard to get work.

### Solution
The basic premise of this script is to check the shared drive
for changes and then when a new Excel file is detected matching
certain naming criteria we obtain its URL, open it in the browser, copy
my name to clipboard, pause all playing music and play a high energy song
to get me hyped to fill my name in quickly!

This python script uses the Microsoft Graph API to authenticate and monitor
for drive changes, then when a change is detected we check that the file
is an Excel file. We also maintain a separate txt file of old rota names
so that when an old rota is modified it is not mistaken for a new file,
since the API endpoint we are using doesn't distinguish between new file
creations and old file modifications. Then we get the URL of the new Excel
file and use the Python webbrowser library to open it in a new tab.
We also make some system calls with the os library to: pause currently
playing media, play music and copy my name to clipboard. This means that once
the rota is opened I can just CTRL-V paste my name and save time by not
having to type it out.

### Requirements

```
pip install oauthlib msal urllib requests webbrowser 
```
