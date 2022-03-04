import scanner
import autofiller
import time
import vlc
import os
import json
import pyautogui as pa
from datetime import datetime as dt


def main(shift_list, drive='personal', relative_path='February 2022',
         play_music=True, afk_mode=True, sleep_time=2):
    """Instantiate both autofiller and scanner objects and combine in a loop."""
    with open('credentials.json', 'r') as f:  # Read in our credentials json
        credentials = json.load(f)
    scope = ['User.Read', 'Files.ReadWrite.All', 'Files.Read.All',
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

    gc.get_access_token()  # Web app client authentication

    with open('old_rotas.txt', 'r') as f:  # Read in old rota names as a list
        old_rotas = f.read().splitlines()

    screen_region = (0, 0, 1080, 1920)  # Docked
    # screen_region = (0,0,2560,1600) # Laptop, not really working
    colours = [(146, 208, 80), (248, 203, 173),
               (68, 114, 196), (0, 0, 0)]  # Cell colours
    # Instantiate an Autofill object
    af = autofiller.Autofill(screen_region=screen_region, colours=colours)

    counter = 0
    num_rotas_autofilled = 0
    start_time = dt.now()  # Keep track of time elapsed
    finished = False
    error_log = []
    while not finished:
        try:
            counter += 1

            mouse_position_before_sleep = pa.position()
            time.sleep(sleep_time)
            mouse_position_after_sleep = pa.position()

            if counter % 300 == 0:
                print("")
                print("Refreshing access token...")
                print("")
                gc.refresh_access_token()  # We don't need to refresh every single loop
                # Detect whether the user is afk
                # If the x-coord of the mouse doesn't change after sleeping
                # then only move the y-coordinate of the mouse
                print("Checking for AFK...")
                print("")
                if mouse_position_before_sleep[0] == mouse_position_after_sleep[0]:
                    afk_mode = True
                    print("AFK mode activated!")
                    print("")

            if mouse_position_before_sleep[0] != mouse_position_after_sleep[0]:
                afk_mode = False  # Deactivate AFK mode

            if afk_mode:
                # Move the cursor vertically down the screen
                pa.moveTo((500, (30*counter % 700)+100))


            drive_items_response = gc.get_driveItems(relative_path)
            time_elapsed = str(dt.now() - start_time).split('.')[0]
            try:
                drive_items = drive_items_response['value']
                print("--------------------------")
                print(f"Scanning {drive} drive...")
                print(f"Check: {counter}")
                print(f"Refresh time: {sleep_time} secs")
                print(f"AFK: {afk_mode}")
                print(f"Time elapsed: {time_elapsed}")
                print(f"Errors: {error_log}")
                print(f"Rotas autofilled: {num_rotas_autofilled}")
                rotas = [(item['name'], item['webUrl']) for item in drive_items]
                print(f"DriveItems detected: {len(rotas)}")
                for rota in rotas:
                    print(rota[0])
                    rota_name, rota_url = rota
                    # Check file is not an old rota and also is an excel spreadsheet
                    if rota_name not in old_rotas and rota_name.split('.')[-1] == 'xlsx':
                        print(f"New rota detected: {rota_name}")
                        if play_music:  # Only trigger once
                            try:
                                # Pause any currently playing media
                                os.system('playerctl stop')
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
                        af.autofill_shifts(
                            shift_list=shift_list, rota_url=rota_url)
                        with open('old_rotas.txt', 'a') as f:
                            # Append old rotas with new rota
                            f.write(rota_name + '\n')
                        print(f"Appending {rota_name} to old_rotas.txt...")
                        old_rotas.append(rota_name)
                        # finished = True
                        num_rotas_autofilled += 1

            except KeyError:
                print("--------------------------")
                print(f"Scanning {drive} drive...")
                print(f"Check: {counter}")
                print(f"Refresh time: {sleep_time} secs")
                print(f"AFK: {afk_mode}")
                print(f"Time elapsed: {time_elapsed}")
                print(f"Errors: {error_log}")
                print(f"Rotas autofilled: {num_rotas_autofilled}")
                print(f"{relative_path} {drive_items_response['error']['code']}")
                print(drive_items_response['error']['message'])

        except KeyboardInterrupt:  # Toggle Away From Keyboard mode using ctrl-c
            print("Keyboard interrupt detected!")
            if afk_mode:
                afk_mode = False
                print("AFK mode off!")
            else:
                afk_mode = True
                print("AFK mode on!")

        except Exception as e:
            print(e)
            print("")
            print("Error occured...")
            print("Appending Error Log...")

            with open('error_log.txt', 'a') as f:
                f.write(str(dt.now()) + ': ' + str(e) + '\n')

            print("")
            print("Sleeping...")
            error_log.append(type(e).__name__)
            time.sleep(10)
            # sleep_time = 10
            print("Retrying...")

    print("")
    print("FINISHED.")
    print("")


if __name__ == "__main__":
    shift_list = [
        ['Isaac Lee', ['sunday morning', 'saturday morning']],
        ['Osaruese Egharevba', ['sunday morning', 'saturday morning', 'thursday afternoon', 'tuesday afternoon']],
        ['Ayse Zeynep Kamis', ['sunday evening', 'thursday evening']],
        ['Nithil Kennedy', ['thursday evening']],
        ['Dominic Haworth-Staines', ['thursday evening']],
        ['Lucile Villeret', ['wednesday afternoon', 'saturday afternoon']],
        ['Michael Pristin', ['thursday evening']]
    ]
    main(shift_list=shift_list, drive='rota', relative_path='March 2022',
         play_music=True, afk_mode=False, sleep_time=2)
