import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox as tkMessageBox
import win32api
import win32gui
import win32process
import os
from pyWinActivate import win_activate, get_app_list

workdir = os.getcwd()
icons_dir = workdir + "\\icons\\"

# get screen working space without taskbar reserved space
monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0,0)))
work_area = monitor_info.get("Work")
scr_width = win32api.GetSystemMetrics(0)
scr_height = win32api.GetSystemMetrics(1) - (win32api.GetSystemMetrics(1) % work_area[3]) # shorten window if taskbar is on the bottom of the screen




monitor_count = len(win32api.EnumDisplayMonitors()) #check how many monitors are plugged in
scr_w_placement = scr_width*monitor_count-scr_width-(win32api.GetSystemMetrics(0) % work_area[2]) # placement of the window based on number of monitors


wmx = None


# CCT opened
# coords 1
# coords 2
# 1 = not ready to start data extraction
# 0 = can start script
start_check = [1,1,1]
override = False

# index 0,1 = country dropdown menu
# index 2,3 = first data row
coords = [0,0,0,0]

#colors
dark_blue = "#00134d"


window_title = "Distribution Panel CCT @GfK"
finished = "Alert! CCT data downloader"

# Close after 1 cycle prompt
prompt = False


def load_coords():
    try:
        file = open("coordinates.txt", "r").readlines()
        for i, line in enumerate(file):
           line = line.split("\n")
           coords[i] = (int(line[0]))
    except FileNotFoundError:
        print("File \"coordinates.txt\" not found.")

def change_coords():
    file = open("coordinates.txt", "w")
    for i in range(4):
        file.write(str(coords[i])+"\n")
    file.close()

def check_CCT_availability():
        for proc in get_app_list():
            if window_title in proc[1]:
                return True

def check_coordinates():
    sums = sum(coords)
    
    if sums == 0:
        start_check[1] = 1
        start_check[2] = 1
    else:
        start_check[1] = 0
        start_check[2] = 0

def load_cycles():
    global prompt
    try:
        file = open("cycles.txt", "r").readlines()
        for line in file:
            line = line.split("\n")
            if line[0] == "True":
                prompt = True
            elif line[0] == "False":
                prompt = False
    except FileNotFoundError:
        print("File \"cycles.txt\" not found.")

def save_cycles():
    file = open("cycles.txt", "w")
    file.write(str(prompt)+"\n")
    file.close()


def get_mouse_pos(event, set_button):
    global wmx, mposxy
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

class UI(tk.Frame):
    def __init__(self, parent):
        global prompt

        tk.Frame.__init__(self, parent)
        backgr = tk.Frame(self, borderwidth=1, relief="sunken")
        backgr.pack(side="bottom", fill="both", expand=True)
        

        self.mainframe = tk.Frame(self, bg=dark_blue)
        self.mainframe.place(relx=0.015, rely=0.015, relwidth=0.97, relheight=0.7)

        self.dpc = tk.Label(self.mainframe, text="Distribution Panel CCT")
        self.dpc.grid(row=0, column=0)

        self.photo0 = tk.PhotoImage(file = icons_dir+"not_found.png")
        self.photo1 = tk.PhotoImage(file = icons_dir+"found.png")
        self.window_found = tk.Label(self.mainframe, image = self.photo0)
        self.window_found.grid(row=0, column=1, pady=10)
        
        
        tk.Label(self.mainframe, text="Country").grid(row=1, column=0, sticky="nw", pady=10)
        tk.Label(self.mainframe, text="First Record Row").grid(row=2, column=0, sticky="nw", pady=10)


        self.coordsdrop = tk.Label(self.mainframe)
        self.coordsdrop.grid(row=1, column=1, sticky="nw", pady=10)

        tk.Button(self.mainframe, text="Change", command=lambda set_button="drop": self.set_coords(set_button)).grid(row=1, column=2, sticky="nw", pady=10, padx=10)
        tk.Button(self.mainframe, text="Check").grid(row=1, column=3, sticky="nw", pady=10)

        self.coordsrow = tk.Label(self.mainframe)
        self.coordsrow.grid(row=2, column=1, sticky="nw", pady=10)

        tk.Button(self.mainframe, text="Change", command=lambda set_button="row": self.set_coords(set_button)).grid(row=2, column=2, sticky="nw", pady=10, padx=10)
        tk.Button(self.mainframe, text="Check").grid(row=2, column=3, sticky="nw", pady=10)

        self.start_button = tk.Button(self, text="Start", command=self.start_script)
        self.start_button.place(relx=0.015, rely=0.68, relwidth=0.97, relheight=0.30)


        # Close program after 1 cycle
        self.what = tk.Button(self.mainframe, text="Close the program after 1 cycle?", command=self.cycles)
        self.what.grid(row=3, column=0, columnspan=2, sticky="nw", pady=10)
        self.cycle = tk.Label(self.mainframe, text=str(prompt))
        self.cycle.grid(row=3, column=3, sticky="nw", pady=10)

        



#==========TIMER===================
        #timer object
        self.loop = tk.Label(root)
        self.loop.place(x=0,y=0,width=0,height=0)
        # start the timer
        self.loop.after(100, self.timer)

    def timer(self):
        global override

        window = check_CCT_availability()
        check_coordinates()

        self.coordsdrop.configure(text="x=%i ,y=%i" % (coords[0], coords[1]))
        self.coordsrow.configure(text="x=%i ,y=%i" % (coords[2], coords[3]))

        if override:
            for proc in get_app_list():
                if finished in proc[1]:
                    override = False
                    self.start_button.configure(text="Start", state=tk.NORMAL)
                    if prompt:
                        raise SystemExit

        if window:
            self.window_found.configure(image = self.photo1)
            start_check[0] = 0
        else:
            self.window_found.configure(image = self.photo0)
            start_check[0] = 1

        start_sum = sum(start_check)
        if override == False:
            if start_sum == 0:
                self.start_button.configure(state=tk.NORMAL)
            else:
                self.start_button.configure(state=tk.DISABLED)
        else:
            pass

        self.loop.after(1000, self.timer)


    def set_coords(self, set_button):
        CoordsSet(root, set_button)
        
    def start_script(self):
        global override

        override = True
        self.start_button.configure(text="Running... Please Wait.", state=tk.DISABLED)
        os.startfile("CTMextractor.pyw")


    def cycles(self):
        global prompt
        prompt = not prompt
        self.cycle.configure(text=str(prompt))
        save_cycles()


class CoordsSet():
    def __init__(self, parent, set_button):
        self.top = tk.Toplevel()
        self.top.title("lines test")
        self.top.wm_attributes("-topmost", 1)
        self.top.attributes("-alpha", 0.5)
        self.top.attributes("-fullscreen", True)
        # self.top.geometry("%ix%i+%i+%i" % (300, 300, 20, 20)) #WidthxHeight and x+y of main window
        self.top.geometry("%ix%i+%i+%i" % (scr_width, scr_height, scr_w_placement, 0)) #WidthxHeight and x+y of main window
        self.top.overrideredirect(True) # removes title bar
        self.top.focus_force()
        self.top.bind("<Key-Escape>", self.close_top)


        self.backgr = tk.Canvas(self.top, background="white")

        self.linex = self.backgr.create_line(0, 0, 0, 0, width=1)
        self.liney = self.backgr.create_line(0, 0, 0, 0, width=1)

        self.backgr.pack(side="bottom", fill="both", expand=True)
        self.backgr.bind("<Button-1>", lambda event, set_button=set_button : get_mouse_pos(event, set_button))
        self.backgr.bind("<Motion>", self.crosshair)
        


        tk.Label(self.backgr, text="Press \"Esc\" to close the coordinates window", bg="white" ,foreground="red", font=("Helvetica", 50)).place(relx=0.15, rely=0.85)
        
        





    def crosshair(self, event):
        self.mpos = win32gui.GetCursorInfo()
        self.mposxy = self.mpos[2]
        self.newcoordsx = (self.mposxy[0], self.mposxy[1]-50, self.mposxy[0], self.mposxy[1]+50)
        self.backgr.coords(self.linex, self.newcoordsx)
        
        self.newcoordsy = (self.mposxy[0]-50, self.mposxy[1], self.mposxy[0]+50, self.mposxy[1])
        self.backgr.coords(self.liney, self.newcoordsy)
        

    def close_top(self, event):
        self.top.destroy()

if __name__ == '__main__':
    root = tk.Tk()   
    root.title("Main Menu")

    root.option_add('*font', ("Helvetica", 15))
    root.option_add('*foreground', ("white"))
    root.option_add('*background', ("#00134d"))
    
    root.configure(background="yellow")
    root.wm_attributes("-topmost", 1)
    root.wm_attributes("-transparentcolor", "yellow")
    root.attributes("-alpha", 0.95)
    root.geometry("%ix%i+%i+%i" % (scr_width/3.5, scr_height/3, scr_width/2, scr_height/3)) #WidthxHeight and x+y of main window

    load_cycles()
    load_coords()

    ui = UI(root)
    ui.place(relwidth=1, relheight=1)
    
    root.mainloop()