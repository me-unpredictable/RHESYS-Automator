import sys
import time
import tkinter as tk
from tkinter import filedialog as of

import keyboard
import pyautogui as ui
import keyboard as kb # for the keypress use 'send' function otherwise u need to use press and release in pair
import os
import pandas as pd



#variables
title='Oracle Fusion Middleware Forms Services - WebUtil - I' # full name is internet explorer, but when
w_title= 'RHESYS Automator v1.0' # This variable contains windows title
uid=pwd=''
rnum=0 # record number
delay=3

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
def erase_txt():
    ui.sleep(2)
    kb.send('end')
    ui.sleep(0.5)
    kb.send('shift+home')
    ui.sleep(0.2)
    kb.send('backspace')
    ui.sleep(0.2)
    #kb.send('backspace')
def fill_prj_tile(data): # this function fills information in project sub tab
    print(data['Project title'][rnum])
    show_win(title)
    ui.scroll(1000)
    loc=ui.locateOnScreen('img/pr_title.png')
    ui.click(loc[0]+100,loc[1]+10)
    kb.write(data['Project title'][rnum])
def fill_prj_des(data): # this function fills information in project sub tab
    print(data['Project title'][rnum])
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/pr_des.png')
    ui.click(loc[0] + 100, loc[1] + 10)
    kb.write(data['Project description'][rnum])
def fill_start(data): # this function fills information in project sub tab
    print(data['Start date'][rnum])
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/pr_start.png')
    ui.click(loc[0] + 100, loc[1] + 10)
    sd=str(data['Start date'][rnum]).split(' ')[0]
    # convert date from yymmdd to ddmmyy
    sd=sd.split('-')
    sd=sd[2]+'/'+sd[1]+'/'+sd[0]
    kb.write(sd)
def fill_fin(data): # this function fills information in project sub tab
    print(data['Project title'][rnum])
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/pr_fin.png')
    ui.click(loc[0] + 100, loc[1] + 10)
    ed = str(data['End date'][rnum]).split(' ')[0]
    # convert date from yymmdd to ddmmyy
    ed = ed.split('-')
    ed = ed[2] + '/' + ed[1] + '/' + ed[0]
    kb.write(ed)
def fill_res(data):
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/res_per.png')
    ui.click(loc[0] + 100, loc[1] + 10)
    kb.write(str(data['% Research'][rnum]))
def fill_overheads(data):
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/levy.png')
    ui.click(loc[0] + 100, loc[1] + 10)
    kb.write(str(data['Overheads'][rnum]))
def fill_ip_owner(data):
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/ip_owner.png')
    ui.click(loc[0] + 100, loc[1] + 10)
    kb.write(str(data['Project IP arrangement'][rnum]))
def fill_rdo_bdo(data): #need to work (drop down)
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/rdo_bdo.png')
    kb.write(str(data['BD contact'][rnum]))
    kb.send('tab')
def fill_cheif_in(data,btn): # double check everyting (it takes more button as argument)
    r_names = str(data['Researcher name'][rnum]).split(';')
    r_names = ''.join(r_names).split('#')
    r_name=[]
    for i in r_names:
        if not (i.isnumeric()):
            r_name.append(i) # add all researcher names to r_name
    # check if there are more than one researchers
    multi=False # flag for multiple researcher (initially we assume that there is only one researcher)
    if len(r_name)>1:
        multi=True
    # add first researcher
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/c_cin.png') # find field
    ui.click(loc[0] + 10, loc[1] + 20) # click on it

    # erase if anything is written (here we can not use erase text function this field is different)
    ui.sleep(2)
    kb.write('123')
    kb.send('enter')
    # now everything is erased
    kb.write('%')
    kb.send('enter') # search in full list
    ui.sleep(delay)
    loc = ui.locateOnScreen('img/find.png')  # find field
    ui.click(loc[0] + 100, loc[1] + 10)  # click on it
    erase_txt()

    name=r_name[0].split(' ') # divide first name and last name
    kb.write(name[1]+'%'+name[0]) # write lastname % first name
    kb.send('enter')
    ui.sleep(1)
    kb.send('enter')


    # Click on more if more researchers exits
    if multi:
        btn['state']='normal'
        ui.sleep(delay)
        # save recored
        ui_save()
        ui.sleep(delay)
        # click on more
        ui.scroll(1000)
        loc = ui.locateOnScreen('img/btn_more.png')  # find field
        ui.click(loc[0] + 10, loc[1] + 10)  # click on it
        ui.alert('There are more than one researchers!!\nClick on Fill More once more researcher window is opened.',w_title)
def fill_more(data,btn):
    show_win(title)
    btn['state']='disabled' # once fill more is used disable it again


def fill_project(data,btn):
    #fill_prj_tile(data)
    #fill_prj_des(data)
    #fill_start(data)
    #fill_fin(data)
    #fill_res(data)
    #fill_overheads(data)
    #fill_ip_owner(data)
    #fill_rdo_bdo(data)
    fill_cheif_in(data,btn)
    #fill_start_date()
#--------------------------------------------------------------
def next_record(lbl): # This function will read next record
    global rnum
    rnum+=1
    lbl.update()
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
    # login window creatio
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
    global rnum
    #rnum=0
    # creation of control window
    control=tk.Tk()
    control.title(w_title)
    control.geometry("+%s"%(s_width-250)+"-%s"%(s_height-380))
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
    #c_record=tk.Label(text='Current Record #%d'%rnum).grid(row=0,column=1,)

    lbl_frame1=tk.LabelFrame(control,text='Navigation')
    lbl_frame1.grid(row=0,column=0)

    btn1=tk.Button(lbl_frame1,text='New Project',command=new_project).grid(row=0,column=0)
    btn_next=tk.Button(lbl_frame1,text='Next Record',command=lambda: next_record(control)).grid(row=1,column=0)

    lbl_frame2=tk.LabelFrame(control,text='Fill Data')
    lbl_frame2.grid(row=1,column=0)
    btn_fill_more = tk.Button(lbl_frame2,state='disabled', text='Fill More', command=lambda: fill_more(data,btn_fill_more))
    btn_fill_more.grid(row=0, column=1) # here we need to put this button separately into grid because we are using it
    # as an argument
    btn_fill_pro=tk.Button(lbl_frame2,text='Fill Project',command=lambda: fill_project(data,btn_fill_more)).grid(row=0,column=0)


    #btn_title = tk.Button(lbl_frame2, text='Prj title', command=lambda: fill_prj_tile(data)).grid(row=0, column=0)
    #btn_des= tk.Button(lbl_frame2, text='Prj Description', command=lambda: fill_prj_des(data)).grid(row=0, column=1)
    #btn_start = tk.Button(lbl_frame2, text='Prj start', command=lambda: fill_start(data)).grid(row=1, column=0)
    #btn_finish = tk.Button(lbl_frame2, text='Prj finish', command=lambda: fill_fin(data)).grid(row=1, column=1)
    lbl_frame3=tk.LabelFrame(control,text='Form Controls')
    lbl_frame3.grid(row=2,column=0)
    btn_more=tk.Button(lbl_frame3,text='More',command=ui_more).grid(row=0,column=0)
    btn_c_cmnt = tk.Button(lbl_frame3,text='Contract Comments', command=ui_c_cmnt).grid(row=0, column=1)
    btn_grants = tk.Button(lbl_frame3,text='Grants', command=ui_grants).grid(row=1, column=0)
    btn_cmnt = tk.Button(lbl_frame3,text='Comments', command=ui_comments).grid(row=1, column=1)
    btn_ua = tk.Button(lbl_frame3,text='User Actions', command=ui_ua).grid(row=2, column=0)
    btn_sta = tk.Button(lbl_frame3,text='Status', command=ui_status).grid(row=2, column=1)
    btn_forms = tk.Button(lbl_frame3,text='Forms', command=ui_forms).grid(row=3, column=0)
    btn_kw = tk.Button(lbl_frame3,text='Keywords', command=ui_keywords).grid(row=3, column=1)
    btn_links = tk.Button(lbl_frame3,text='Links', command=ui_links).grid(row=4, column=0)
    btn_save = tk.Button(lbl_frame3,text='Save', command=ui_save).grid(row=4, column=1)
    btn_exit = tk.Button(lbl_frame3,text='Exit Form', command=ui_exit).grid(row=5, column=1)
#    settings_button=tk.Button(text='Settings',command)
    exit_btn=tk.Button(text='Close',command=shut).grid(row=4,column=1)
    # configuration to keep control window always on top
    control.attributes('-topmost',True)
    #control.update()
    control.mainloop()
    #print('Good bye')

#login_win()
#get_data()
#print(kb._canonical_names.canonical_names)
main_win()
#s_test(-1000)
