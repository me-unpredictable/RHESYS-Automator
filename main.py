import time
import tkinter as tk
import pyautogui as ui
import os


#wb=webdriver.Ie('IEDriverServer.exe')

import tkinter.ttk

#variables
title='Oracle Fusion Middleware Forms Services - WebUtil - Internet Explorer'
s_width=0
s_height=0

# ----------------------------------------------
# main ui functions
# ----------------------------------------------
def launch_ie(): #function to launch internet explorer
    # note: if internet explorer is launched using this function then it will be closed when clos this script/program
    # currently it is unused
    url = "https://researchuat.uws.edu.au/forms/frmservlet?config=rmsmenu"
    ppath = os.environ['PROGRAMFILES'].split('\\')
    ppath = ppath[0] + '\\' + ppath[1] + '\Internet Explorer\iexplore.exe'

    command = "\"%s\" " % ppath + ' %s' % url
    os.system(command)

def login_(uid,pwd): # this function is to perfrom login
    # get cursor to first field (safety)
    ui.hotkey('ctrl','up')
    # type user id
    ui.typewrite(uid)
    # move to next field
    ui.hotkey('tab')
    # type password
    ui.typewrite(pwd)
    # login
    ui.hotkey('tab')

def check(win):
    titles = ui.getAllTitles()
    # check if orcale fussion is opened or not
    if title not in titles:
        print(titles)
        ui.alert('Please try again after opening RHESYS page in internet explorer')
    else:
        # check if id and password are entered or not
        if uid!='' and pwd!='':
            tk.Tk.destroy(win)
        else:
            ui.alert('Please Enter User Id and Password')
def show_win(title): # function to activate ie window
    # this function needs to be called everytime after clicking on any button
    # it will activate ie window
    ui.getWindowsWithTitle(title)[0].activate()
    ui.getWindowsWithTitle(title)[0].maximize()

def switch(title,uid,pwd): # this function is to bring RHESYS windows in forground
    ui.getWindowsWithTitle(title)[0].activate()
    ui.getWindowsWithTitle(title)[0].maximize()
    # wait for 1 second
    time.sleep(2)
    # call login function
    login_(uid,pwd)

# ----------------------------------------------
# navigation function
# ----------------------------------------------

def new_project(): # thi function is to create new project entry
    show_win(title) # get ie window
    ui.hotkey('alt')
    ui.hotkey('p')
    ui.sleep(0.2)
    ui.hotkey('p')



# ----------------------------------------------
# windows/ gui
# ----------------------------------------------

def login_win():
    global uid
    global pwd
    # login window creation
    login=tk.Tk()
    login.title('Login. . .')
    login.resizable(0,0)
    #get screen size
    s_width=login.winfo_screenwidth()
    s_height=login.winfo_screenheight()
    print(s_width,s_height)

    login.geometry("+%s"%(s_width-110)+"-%s"%(s_height-130))
    login.protocol('WM_DELETE_WINDOW',lambda : ui.alert('Click on exit'))
    #tk variables
    u_id=tk.StringVar()
    p_wd=tk.StringVar()

    uid=tk.Entry(login,textvariable=u_id,width=15).grid(row=0,column=1)
    pwd=tk.Entry(login,textvariable=p_wd,show='*',width=15).grid(row=1,column=1)
    login_btn=tk.Button(login,text='Login',command=lambda :check(login)).grid(row=2,column=1)
    exit_btn=tk.Button(login,text='Exit',command=lambda :exit(0)).grid(row=3,column=1)
    tk.mainloop() # main loop for login window

    uid=u_id.get()
    pwd=p_wd.get()
    switch(title,uid,pwd)
#-------------------------------------
def main_win():
    # creation of control window
    control=tk.Tk()
    control.title('Testing')
    control.geometry("+%s"%(s_width-110)+"-%s"%(s_height-130))
    control.protocol('WM_DELETE_WINDOW',lambda : ui.alert('Click on exit'))
    control.resizable(0,0)
    #space=tk.Label().grid(row=0,column=1,padx=15)
    btn1=tk.Button(text='Switch',command=lambda :switch(title,uid,pwd)).grid(row=1,column=1)
    btn2=tk.Button(text='New',command= new_project).grid(row=2,column=1)
    exit_btn=tk.Button(control,text='Exit',command=lambda :exit(0)).grid(row=3,column=1)
    # configuration to keep control window always on top
    control.attributes('-topmost',True)
    control.update()
    tk.mainloop()
    print('Good bye')

login_win()