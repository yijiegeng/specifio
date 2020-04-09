from receiver import *
import time

if __name__ == "__main__":
    try:
        print("Server start...")
        while True:
            status = process_file()

            if status == 'success':
                print('\n' + "Process done, sleeping for 30 seconds..." + '\n')
                time.sleep(30)
            else:
                print("Sleeping for 30 seconds..." + '\n')
                time.sleep(30)
    except:
        print('\n' + "Server disconnected" + '\n')
