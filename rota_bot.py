import rota_scanner
import rota_autofiller
import time
import vlc
import os

def main():
    shift_list = [['Isaac Lee', ['wednesday evening', 'saturday evening', 'sunday morning']],
                  ['Osaruese Egharevba', ['tuesday evening', 'wednesday afternoon', 'thursday evening', 'saturday morning', 'sunday morning']],
                  ['Nithil Kennedy', ['thursday evening']]]
    rota = rota_scanner.scan_for_new_rota(root_drive='personal')
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
            rota_autofiller.autofill_rota(shift_list=shift_list)

            with open('old_rotas.txt', 'a') as f:
                f.write(rota) # Append old rotas with new rota
            print(f"Appending {rota} to old_rotas.txt...")
        counter += 1
        print("Finished!")

if __name__ == "__main__":
    main()
