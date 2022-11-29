#from ast import Global
#from cProfile import label
import tkinter as tk
#from tkinter import simpledialog
#from tkinter import messagebox as tkMessageBox
#import win32api
import win32gui
#import win32process
#import os
#from pyWinActivate import win_activate, get_app_list
from CTMextractor import ctm
import keyboard
#import time
import threading
#import sys
import pyautogui



# index 0,1 = country dropdown menu
# index 2,3 = first data row
coords = [0,0,0,0]

finished = "Alert! CCT data downloader"

#colors
dark_blue = "#00134d"
dark_gray = "#323f54"
purple = "#7d25b0"
dirty_white = "#ebe8ed"


keyboard.add_hotkey("ctrl+alt+b", lambda: set_mouse_pos("drop"))
keyboard.add_hotkey("ctrl+alt+m", lambda: set_mouse_pos("row"))


def set_mouse_pos(set_button):
    global mposxy
    # mouse pos on screen
    try:
        mpos = win32gui.GetCursorInfo()
        mposxy = mpos[2]
        # wmx = mposxy[0]

        if set_button == "drop":
            coords[0] = mposxy[0]
            coords[1] = mposxy[1]
        elif set_button == "row":
            coords[2] = mposxy[0]
            coords[3] = mposxy[1]
    except:
        print("No mouse coords found.")

    change_coords()

def load_coords():
    try:
        file = open("coordinates.txt", "r").readlines()
        for i, line in enumerate(file):
           line = line.split("\n")
           coords[i] = (int(line[0]))
    except FileNotFoundError:
        print("File \"coordinates.txt\" not found.")

    coordsdrop.configure(text="x=%i ,y=%i" % (coords[0], coords[1]))
    coordsrow.configure(text="x=%i ,y=%i" % (coords[2], coords[3]))

def change_coords():
    file = open("coordinates.txt", "w")
    for i in range(4):
        file.write(str(coords[i])+"\n")
    file.close()


    coordsdrop.configure(text="x=%i ,y=%i" % (coords[0], coords[1]))
    coordsrow.configure(text="x=%i ,y=%i" % (coords[2], coords[3]))


class Window(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        
        global coordsdrop, coordsrow, start_button, mouseposlbl, num_of_loops

        label_frame = tk.LabelFrame(master, text = "Countries (ctrl+alt+b)")
        label_frame.pack(fill=tk.BOTH, padx=5, pady=10)

        coordsdrop = tk.Label(label_frame, text="x0 - y0")
        coordsdrop.pack(side=tk.RIGHT)



        label_frame = tk.LabelFrame(master, text = "First Row (ctrl+alt+m)")
        label_frame.pack(fill=tk.BOTH, padx=5, pady=10)

        coordsrow = tk.Label(label_frame, text="x0 - y0")
        coordsrow.pack(side=tk.RIGHT)

        label_frame = tk.LabelFrame(master, text = "Active mouse position")
        label_frame.pack(fill=tk.BOTH, padx=5, pady=10)

        mouseposlbl = tk.Label(label_frame, text="xy")
        mouseposlbl.pack(side=tk.RIGHT)


        num_of_loops = tk.Entry(master, width = 2)
        num_of_loops.insert(0, 73)
        num_of_loops.pack(side=tk.RIGHT, padx = 5, pady = 5)


        start_button = tk.Button(master, text="Start", bg="green", command=self.start_script)
        start_button.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)



    def start_script(self):
        global process
        start_button.configure(text="Running... Please Wait.", bg="gray", state=tk.DISABLED)


        nloops = int(num_of_loops.get())

        # ctm(x1=coords[0], y1=coords[1], x2=coords[2], y2=coords[3]))



        try:
            print("Starting threaded process...")
            process = threading.Thread(target=ctm, args=(coords[0], coords[1], coords[2], coords[3], nloops))
            process.start()
        except Exception as e:
            print("Threaded process start failed.")
            print(e)
            pyautogui.alert(text='Threaded process start failed.\n{e}', title='CCT data downloader ERROR!', button='OK')


def looped_task():
    #check if extractor process is alive, and set start button accordingly
    try:
        if not process.is_alive():
            start_button.configure(text="Start", bg="green", fg="white", state=tk.NORMAL)
            print("Threaded process not found.")
    except: pass


    mpos = win32gui.GetCursorInfo()
    mposxy = mpos[2]
    mouseposlbl.configure(text="x%i, y%i" % (mposxy[0], mposxy[1]))

    root.after(2000, looped_task)  # reschedule event in 2 seconds


if __name__ == '__main__':
    root = tk.Tk()   
    root.title("CTM Data Downloader")
    root.option_add('*font', ("Helvetica", 15))
    root.option_add('*foreground', ("white"))
    root.option_add('*background', (dark_gray))
    root.configure(background=dark_gray)
    root.wm_attributes("-topmost", 1)
    root.geometry("350x300+750+300") #WidthxHeight and x+y of main window
    root.attributes('-toolwindow', True)
    root.resizable(0,0)
    
    window = Window(root)
    root.after(2000, looped_task)
    load_coords()


    root.mainloop()


