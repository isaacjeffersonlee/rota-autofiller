# Rota Bot

#### Date: 15/01/2022

### Demo
![Demo Example](./demo/rota_bot_demo.GIF)

### Problem
I work for my University's student bar and we have over 100
employees. Each Monday a new rota is released on Microsoft Sharepoint
and you fill your name next to the shift you want.
With the large number of employees, the shifts get taken extremely
quickly so it is often hard to get work.

### Solution
The basic premise of this script is to check the shared drive
for new Excel files and then when a new file is detected matching
certain naming criteria we obtain its URL, open it in the browser and autofill
the shifts with the correct name(s) by the desired times.

This python script uses the Microsoft Graph API to authenticate and scan the shared
drive for new files. Then when a new file is detected we check that the file
is an Excel file and get the URL of the new file.
Then we open it in a browser using Python webbrowser library.
Then once the rota is open in the browser, the autofiller engages and 
uses the PyAutogui library to scan the screen and get the exact pixel coordinates
of the centre of each cell in each shift.
This is accomplished by taking a screenshot of the screen and examining the different
pixel colours, since each shift is a different colour.
Finally according to a predefined list of desired shifts, each persons name is written in the cell corresponding to
the desired time and day.  

An example of the list of desired shifts we can pass:
```
shift_list = [
        ['Isaac Lee', ['sunday morning', 'saturday evening']],
        ['Osaruese Egharevba', ['sunday morning', 'monday morning', 'tuesday evening', 'thursday morning', 'thursday evening', 'saturday morning']],
        ['Nithil Kennedy', ['thursday evening']]
        ]
```

### Requirements

```
pip install numpy oauthlib msal urllib requests webbrowser 
```
