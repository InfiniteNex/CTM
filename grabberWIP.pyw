import keyboard
import win32gui
import os
import time
import pyautogui
from pyWinActivate import win_activate, win_wait_active, get_app_list


""" ALL POSSIBLE WINDOWS'S TITLES
Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\\\Remote
Distribution Panel CCT *ERROR - \\\\Remote
TDistribution Panel CCT *WARNING - \\\\Remote
Loading... - \\\\Remote
"""


title = "Distribution Panel CCT @GfK - [dip.gfk.com]"
no_records_found = "TDistribution Panel CCT *WARNING - \\\\Remote"
change_country_error = "Distribution Panel CCT *ERROR - \\\\Remote"
total_countries = 33

def sleep(sec=0.5):
    time.sleep(sec)

def press(key=str()):
    keyboard.press_and_release(key)

def get_country_data():
    change_country()
    
    try:
        error_check(title=no_records_found)
    except:
        try:
            error_check(title=change_country_error)
        except:
            pass
        copy_data()

def change_country():
    pyautogui.click(x=332 ,y=114)
    sleep()
    press('down')
    sleep()
    press('enter')
    sleep()
    press('tab')
    sleep()
    press('tab')
    sleep()
    press('tab')
    sleep()
    press('enter')
    sleep()
    win_wait_active(win_to_wait=title, exception=no_records_found, message=False)
    sleep()

    # check and clear error if it exists
    try:
        error_check(title=change_country_error)
    except:
        print("No error window found. No actions required.")

    #focus back on main window
    win_activate(window_title=title, partial_match=True)
    win_wait_active(win_to_wait=title, message=False)

def copy_data():
    sleep()
    pyautogui.click(x=43 ,y=205)
    sleep()
    keyboard.send('ctrl+a')
    sleep()
    # copy and add to dataframe
    print("copy and add to dataframe")


def error_check(title=str()):
        win_activate(window_title=title)
        win_wait_active(win_to_wait=title, message=False)
        sleep()
        press('esc')


'''
MAINLOOP
'''
win_activate(window_title=title, partial_match=True)
win_wait_active(win_to_wait=title, message=False)
for i in range(32):
    get_country_data()

