import scanner
import autofiller
import time
import vlc
import os
import json
import pyautogui as pa
from datetime import datetime as dt

def main(shift_list, drive='personal', relative_path='February 2022', play_music=True, afk_mode=True):
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
    # My personal drive id, used for testing
    personal_driveid = 'b!VVp9GUD9v06cyfM41FM4_CwrynnMOI5LpvhML0mbJnIpc8sEVKhASL9LnZIC63dh'
    # Instantiate a GraphClient object
    if drive == 'personal':
        root_driveid = personal_driveid
    elif drive == 'rota':
        root_driveid = rota_driveid
    else:
        print(f"{drive} is not a valid drive!")
        print("Use either 'personal', or 'rota'")

    # Instantiate GraphClient object to authenticate and get driveItems
    gc = scanner.GraphClient(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            scope=scope,
            account_type=account_type,
            root_driveid=root_driveid)

    gc.get_access_token() # Web app client authentication

    with open('old_rotas.txt', 'r') as f: # Read in old rota names as a list
        old_rotas = f.read().splitlines()

    screen_region = (0,0,1080,1920) # Docked
    # screen_region = (0,0,2560,1600) # Laptop, not really working
    colours = [(146,208,80), (248,203,173), (68,114,196), (0,0,0)] # Cell colours
    # Instantiate an Autofill object
    af = autofiller.Autofill(screen_region=screen_region, colours=colours)

    counter = 0
    num_rotas_autofilled = 0
    start_time = dt.now() # Keep track of time elapsed
    sleep_time = 1 # Time to sleep in seconds, setting this too low could cause errors
    finished = False
    while not finished:
        try:
            counter += 1
            time.sleep(sleep_time)
            if afk_mode: pa.moveTo(((10*counter % 500)+100, (10*counter % 500)+100))
            if counter % 30: gc.refresh_access_token() # We don't need to refresh every single loop
            driveItems = gc.get_driveItems(relative_path)['value']
            time_elapsed = str(dt.now() - start_time).split('.')[0]
            print("--------------------------")
            print(f"Scanning {drive} drive...")
            print(f"Check: {counter}")
            print(f"Refresh time: {sleep_time} secs")
            print(f"AFK: {afk_mode}")
            print(f"Time elapsed: {time_elapsed}")
            print(f"No. rotas autofilled: {num_rotas_autofilled}")
            rotas = [(item['name'], item['webUrl']) for item in driveItems]
            print(f"No. items detected: {len(rotas)}")
            for rota in rotas:
                print(rota[0])
                rota_name, rota_url = rota
                # Check file is not an old rota and also is an excel spreadsheet
                if rota_name not in old_rotas and rota_name.split('.')[-1] == 'xlsx':
                    print(f"New rota detected: {rota_name}")
                    if play_music: # Only trigger once
                        try:
                            os.system('playerctl stop') # Pause any currently playing media
                            print("Music paused!")
                        except:
                            print("Error pausing media!")
                        try:
                            song = vlc.MediaPlayer('music.mp3')
                            song.play()
                        except:
                            print("Error playing music!")
                        play_music = False

                    print("Engaging autofiller!")
                    af.autofill_shifts(shift_list=shift_list, rota_url=rota_url)
                    with open('old_rotas.txt', 'a') as f:
                        f.write(rota_name + '\n') # Append old rotas with new rota
                    print(f"Appending {rota_name} to old_rotas.txt...")
                    old_rotas.append(rota_name)
                    # finished = True
                    num_rotas_autofilled += 1

        except KeyboardInterrupt: # Toggle Away From Keyboard mode using ctrl-c
            print("Keyboard interrupt detected!")
            if afk_mode:
                afk_mode = False
                print("AFK mode off!")
            else:
                afk_mode = True
                print("AFK mode on!")
            
        except Exception as e: 
            print(e)
            print("Error occured...")
            print("Sleeping 10 secs...")
            time.sleep(10)
            sleep_time = 10
            print("Retrying...")

    print("")
    print("FINISHED.")
    print("")

if __name__ == "__main__":
    shift_list = [
            ['Isaac Lee', ['sunday morning']],
            ['Osaruese Egharevba', ['sunday morning']],
            ['Isaac Lee', ['saturday evening']],
            ['Osaruese Egharevba', ['monday morning', 'tuesday evening', 'thursday morning', 'thursday evening', 'saturday morning']],
            ['Nithil Kennedy', ['thursday evening']]
            ]
    # Demo Shift List
    # shift_list = [
    #         ['Phil Collins', ['sunday morning', 'saturday morning']],
    #         ['Henry Rollins', ['saturday evening', 'sunday morning', 'sunday evening']],
    #         ['Tracy Chapman', ['sunday afternoon']],
    #         ['Dunununun Batman!', ['sunday afternoon']]
    #         ]
    main(shift_list=shift_list, drive='rota', relative_path='February 2022', play_music=True, afk_mode=True)
