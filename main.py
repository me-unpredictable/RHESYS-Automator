import sys
import time
import tkinter as tk
from tkinter import filedialog as of
import pyautogui as ui
import keyboard as kb
import os
import pandas as pd



#variables
title='Oracle Fusion Middleware Forms Services - WebUtil - I' # full name is internet explorer, but when
w_title= 'RHESYS Automator v1.0' # This variable contains windows title
uid=pwd=''


#citix is ued full name is not shown in citrix Ie. To support citrix title needs to be till I.
# if we end title earlier then it can take anyother browser as main browser
# to make sure we are opening Internet explorer, title has to finish on I.

# get screen size
s_info= tk.Tk()
s_info.update_idletasks()
s_width=s_info.winfo_screenwidth()
s_height=s_info.winfo_screenheight()
s_info.destroy()



# ----------------------------------------------
# main ui functions
# ----------------------------------------------
def shut():
    sys.exit()
    exit(0)
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
    #ui.hotkey('ctrl', 'up')
    kb.send('ctrl+up')
    # type user id
    ui.typewrite(uid)
    ui.sleep(0.1)
    # move to next field
    kb.send('tab')
    ui.sleep(0.1)
    # type password
    ui.typewrite(pwd)
    ui.sleep(0.1)
    # login
    kb.send('tab')

def check(win):
    titles = ui.getAllTitles()
    # check if orcale fussion is opened or not
    err=True # assume that our app is not opened
    for t_list in titles:
        if title in t_list:
            err=False
    if err:
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
    ui.sleep(0.1)
    ui.moveTo(s_width/2,s_height/2)

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
    kb.send('alt')
    kb.send('p')
    ui.sleep(0.3)
    kb.send('p')

def get_data(window,f_path): # This function reads excel file data
    data=pd.read_excel(f_path) # open excel file
    cols=data.shape[1] # get number of columns
    # this feature will prevent user from selecting wrong file
    if cols<81:
        ui.alert('Wrong File, Try again',title=w_title)
        window.destroy() # destroy main window
        main_win() # reopen main window

    return data

def fill_project(): # this function fills information in project sub tab
    prject_title=''

def ui_more(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/btn_more.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_c_cmnt(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(1000)
    loc=ui.locateOnScreen('img/btn_c_comments.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_grants(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_grants.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_comments(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_comments.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_ua(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_user_action.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_status(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_status.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_forms(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_forms.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_keywords(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_keyword.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_links(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_links.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_save(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_save.png')
    ui.click(loc[0]+3,loc[1]+3)
def ui_exit(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_exit.png')
    ui.click(loc[0]+3,loc[1]+3)

# ----------------------------------------------
# windows/ gui
# ----------------------------------------------

def login_win():
    global uid,pwd
    # login window creation
    login=tk.Tk()
    login.title('Login. . .')
    login.resizable(0,0)


    login.geometry("+%s"%(s_width-110)+"-%s"%(s_height-130))
    login.protocol('WM_DELETE_WINDOW',lambda : ui.alert('Click on exit'))
    #tk variables
    u_id=tk.StringVar()
    p_wd=tk.StringVar()

    uid=tk.Entry(login,textvariable=u_id,width=15).grid(row=0,column=1)
    pwd=tk.Entry(login,textvariable=p_wd,show='*',width=15).grid(row=1,column=1)
    login_btn=tk.Button(login,text='Login',command=lambda :check(login)).grid(row=2,column=1)
    exit_btn=tk.Button(login,text='Exit',command=lambda :exit(0)).grid(row=3,column=1)
    uid = u_id.get()
    pwd = p_wd.get()
    switch(title, uid, pwd)
    login.mainloop() # main loop for login window



#-------------------------------------
def main_win():
    rnum=0
    # creation of control window
    control=tk.Tk()
    control.title(w_title)
    control.geometry("+%s"%(s_width-130)+"-%s"%(s_height-450))
    control.protocol('WM_DELETE_WINDOW',lambda : ui.alert('Click on exit'))
    control.resizable(0,0)

    # ask for the file name
    ftypes=(
        ('Excel Workbook', '*.xlsx'),
        ('Excel 97- Excel 2003 Workbook', '*.xls')
    )
    file=of.askopenfilename(title='Select Share point Excel File',
                            filetypes=ftypes
                            )
    if file=='':
        ui.alert('You didn\'t selecte a file, exiting...', title=w_title)
        exit(0)
    data=get_data(control,file)
    #print(data['Project title'][rnum])
    #exit(0)
    c_record=tk.Label(text='Current Record #%d'%rnum).grid(row=0,column=1,)
    btn1=tk.Button(text='New Project',command=new_project).grid(row=1,column=1)
    btn2=tk.Button(text='New',command= new_project).grid(row=2,column=1)
    btn_next=tk.Button(text='Next Record',command=lambda : (sum(rnum,1),print(rnum) )).grid(row=3,column=1)
    btn_more=tk.Button(text='More',command=ui_more).grid(row=4,column=1)
    btn_c_cmnt = tk.Button(text='Contract Comments', command=ui_c_cmnt).grid(row=5, column=1)
    btn_grants = tk.Button(text='Grants', command=ui_grants).grid(row=6, column=1)
    btn_cmnt = tk.Button(text='Comments', command=ui_comments).grid(row=7, column=1)
    btn_ua = tk.Button(text='User Actions', command=ui_ua).grid(row=8, column=1)
    btn_sta = tk.Button(text='Status', command=ui_status).grid(row=9, column=1)
    btn_forms = tk.Button(text='Forms', command=ui_forms).grid(row=10, column=1)
    btn_kw = tk.Button(text='Keywords', command=ui_keywords).grid(row=11, column=1)
    btn_links = tk.Button(text='Links', command=ui_links).grid(row=12, column=1)
    btn_save = tk.Button(text='Save', command=ui_save).grid(row=13, column=1)
    btn_exit = tk.Button(text='Exit Form', command=ui_exit).grid(row=14, column=1)
    exit_btn=tk.Button(text='Close',command=shut).grid(row=15,column=1)
    # configuration to keep control window always on top
    control.attributes('-topmost',True)
    #control.update()
    control.mainloop()
    #print('Good bye')

#login_win()
#get_data()
main_win()
#s_test(-1000)