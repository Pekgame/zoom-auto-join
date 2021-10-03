print("Loading. . .")
import os
import time, sys
from datetime import datetime
import meeting_details
try:
    import pyautogui, openpyxl
    os.system("cls")
except ImportError as e:
    os.system("cls")
    print(f'error, install the requirements:\n{e}') 
    sys.exit(e)
    
ZOOM_PATH = "C:/Users/win10/AppData/Roaming/Zoom/bin/Zoom.exe"

def join_meeting(meeting_id, meeting_passcode):
    """join the meeting by pressing images on screen
    :parm meeting_id: the meeting's id
    :parm meeting_passcode: the meeting's passcode
    :type meeting_id: str
    :type meeting_passcode: str
    """

    os.startfile(ZOOM_PATH)

    # locate the '+' icon to join a meeting
    button = pyautogui.locateOnScreen('images/join_a_meeting.png') 
    while button is None:
        button = pyautogui.locateOnScreen('images/join_a_meeting.png')
        
    pyautogui.click(button)

    button = pyautogui.locateOnScreen('images/exit_box.png') 
    while button is None:
        button = pyautogui.locateOnScreen('images/exit_box.png')
        
    pyautogui.click(button)


    # loctae 'enter meeting link' text box
    button = pyautogui.locateOnScreen('images/enter_link.png')
    while button is None:
        button = pyautogui.locateOnScreen('images/enter_link.png')

    pyautogui.click(button)
    pyautogui.write(meeting_id) # enter meeting id

    # join without audio
    #button = pyautogui.locateCenterOnScreen('images/check_box.png')
    #while button is None:
    #    button = pyautogui.locateCenterOnScreen('images/check_box.png')

    pyautogui.click(button)

    # join without video
    button = pyautogui.locateCenterOnScreen('images/check_box.png')
    while button is None:
        button = pyautogui.locateCenterOnScreen('images/check_box.png')

    pyautogui.click(button)

    # join button in 'enter meeting link' screen
    button = pyautogui.locateOnScreen('images/join_btn.png')
    while button is None:
        button = pyautogui.locateOnScreen('images/join_btn.png')

    pyautogui.click(button)

    # type meeting's passcode if it has one
    if meeting_passcode is not None:
        button = pyautogui.locateOnScreen('images/enter_passcode.png')
        while button is None:
            button = pyautogui.locateOnScreen('images/enter_passcode.png')

        pyautogui.click(button)
        pyautogui.write(meeting_passcode)

        button = pyautogui.locateOnScreen('images/join_btn.png')
        while button is None:
            button = pyautogui.locateOnScreen('images/join_btn.png')
        
        pyautogui.click(button)
        
    button = pyautogui.locateOnScreen('images/zoom_box.png')
    while button is None:
        button = pyautogui.locateOnScreen('images/zoom_box.png')
        
    pyautogui.click(button)
    button = pyautogui.locateOnScreen('images/zoom_box.png')
    while button is None:
        button = pyautogui.locateOnScreen('images/zoom_box.png')
        
    pyautogui.click(button)
    exit(0)

    
def main():

    meetings_start_time = input("\nWhat is meeting start time?(Ex. 08:28 or 13:52):\n")
    os.system("cls")
    use_link = str(input("Did you want to use meeting link?(y/n):\n"))
    os.system("cls")
    if use_link.lower() != "y":
        meeting_id = str(input("Enter meeting id:\n"))
        os.system("cls")
        meeting_passcode = str(input("Enter meeting passcode:\n"))
        os.system("cls")
    else:
        meeting_link = str(input("Enter meeting link:\n"))
        os.system("cls")
        meeting_id = meeting_details.get_meeting_id(meeting_link)
        meeting_passcode = meeting_details.get_meeting_passcode(meeting_link)
    print("Complete setting!\nThe meeting will start soon!")
    while True:
        current_time = datetime.now()
        current_time = current_time.strftime('%H:%M')

        if str(current_time) == str(meetings_start_time):
            os.system("cls")
            print("Joining...")
            join_meeting(meeting_id, meeting_passcode)

            
if __name__ == '__main__':
    print("Welcome to Zoom Auto Join!\nThe app is wait for setting!")
    main()
