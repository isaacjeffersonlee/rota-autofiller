import scanner
import autofiller
import time
import vlc
import os

def main():
    shift_list = [['Isaac Lee', ['wednesday evening', 'saturday evening', 'sunday morning']],
                  ['Osaruese Egharevba', ['tuesday evening', 'wednesday afternoon', 'thursday evening', 'saturday morning', 'sunday morning']],
                  ['Nithil Kennedy', ['thursday evening']]]
    rota = scanner.scan_for_new_rota(root_drive='personal')
    try:
        os.system('playerctl stop') # Pause any currently playing media
        print("Music paused!")
    except:
        print("Error pausing media!")

    counter = 0
    while True: # Continue playing song
        if counter == 0: #Only play on first loop
            try:
                song = vlc.MediaPlayer('music.mp3')
                song.play()
            except:
                print("Error playing music!")

            time.sleep(1)
            print("Engaging Autofiller!")
            screen_region = (0,0,1080,1920)
            colours = [(146,208,80), (248,203,173), (68,114,196), (0,0,0)]
            af = autofiller.Autofill(shift_list=shift_list, screen_region=screen_region, colours=colours)
            af.autofill_shifts()

            with open('old_rotas.txt', 'a') as f:
                f.write(rota) # Append old rotas with new rota
            print(f"Appending {rota} to old_rotas.txt...")
        counter += 1

if __name__ == "__main__":
    main()
