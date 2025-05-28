 #File:     KinClk.py
 #Version:  0.0.01
 #Author:   Susan Haynes
 #Comments/Notes:
 #  (0,0) coordinates are the top left corner of the screen for 1920x1080
 #  (0,0) coordinates are the bottom right corner of the screen for 1919x1079
 #To find the location on a screen open IDLE
 #>>> import pyautogui      <- this allows us to use pyautogui prompts
 #>>> pyautogui.size()    <- this returns the size of the monitor
 #>>> pyautogui.position()  <- this returns the exact location of where the mouse pointer is

import configparser                                     # parsing multiple GUI's
import datetime as dt                                   # Date library
import keyboard                                         # windows right key
import os                                               # closing an executable
import pyautogui                                        # automating screen clicks
import pymem                                            # checking if .exe is open
import pymem.process                                    # checking if .exe is open
import pywinauto                                        # bringing an .exe to the foreground
import subprocess                                       # open an executable
import time                                             # call time to count/pause
import tkinter as tk                                    # Tkinter's Tk class
import tkinter.ttk as ttk                               # Tkinter's Tkk class
import win32con                                         # justify right or left the GUI.
import win32gui                                         # bring apps to front foreground

from functools import partial                           # freezing one function while executing another
from openpyxl import *                                  # Write to excel
from pathlib import PureWindowsPath                     # library that cleans up windows path extensions
from PIL import ImageTk, Image                          # Displaying LAL background photo
from python_imagesearch.imagesearch import imagesearch  # opening images, pip package
from tkinter import messagebox                          # Exit standard message box
from win32gui import GetWindowText, GetForegroundWindow # check position of a window

config = configparser.ConfigParser()
samp_arr = []               # initalize all global variables and global arrays to call between classes
btn_pres_cnt = 1            # setting count to 0 to be able to call it a global variable within the function

###########################################################################################################################################
##################################################    KINESIS & LUMEDICA     ##############################################################
###########################################################################################################################################

def kin_main():                                                             # bring Kinesis to the main screen, we will need this multiple times.
    kin_title = 'Kinesis'
    kin_app = pywinauto.Application().connect(title=kin_title)
    kin_win = kin_app[kin_title]
    kin_win.set_focus()
    print('Kinesis is in the foreground now.')

def kin_pop():                                                              # bring Kinesis POPUP to the main screen, we will need this multiple times.
    pop_title = 'Sequence Options'                                          # DOUBLE CHECK THIS IS THE POPUP WINDOW NAME!!!!!!!!!!!!!!!!!!!!!!!!!!!
    pop_app = pywinauto.Application().connect(title=pop_title)
    pop_win = pop_app[pop_title]
    pop_win.set_focus()
    print('Kinesis Pop-up is in the foreground now.')
    print("Do we want to add click resume in this function?? Why else would we want the pop up in the foreground?")

#def lum_main():                                                           # bring Lumedica to the main screen, we will need this multiple times.
#    lum_title = 'Lumedica'
#    app = pywinauto.Application().connect(title=kin_title)
#    kin_win = app[kin_title]
#    kin_win.set_focus()
#    print('Lumedica is in the foreground now.')

try:                                                                        # Try: check if Kinesis is already open.                
    kin_pm = pymem.Pymem('Kinesis.exe')
    print('Kinesis is already open.')
except:                                                                     # Except, if not open, then open it.
    print('Kinesis is not running, lets open it!!!')
    subprocess.Popen('C:\\Program Files\\Thorlabs\\Kinesis\\Thorlabs.MotionControl.Kinesis.exe', shell=True) # Open Kinesis sw
    time.sleep(8)                                                           # wait 8 seconds for kinesis to open fully

#try:                                                                      # Try: check if Lumedica is already open.                
#    lum_pm = pymem.Pymem('Lumedica.exe')
#    print('Lumedica is already open.')
#except:                                                                   # Exception executed, if not open, then open it.
#    print('Lumedica is not running, lets open it!!!')
#    subprocess.Popen('C:\\Program Files\\Thorlabs\\Kinesis\\Thorlabs.MotionControl.Lumedica.exe', shell=True) # Lumedica sw
#    time.sleep(8)                                                         # wait 8 seconds for Lumedica to open fully

## Right justify Kinesis GUI
kin_main()                                                                  # Bring Kinesis to main foreground
time.sleep(1)                                                               # pause to allow to come to foreground
hwnd = win32gui.GetForegroundWindow()                                       # grab the window in the foreground
rect = win32gui.GetWindowRect(hwnd)                                         # assign window rectangle coordinates to an array
a = rect[0]                                                                 # a=upper left corner positon of the Kinesis window in the X coordinates of the screen
b = rect[1]                                                                 # b=upper left corner positon of the Kinesis window in the Y coordinates of the screen
c = rect[2] - a                                                             # c is the length of the kinesis window, should be half the length of the screen 1920/2=960
d = rect[3] - b                                                             # d is the height of the kinesis window, should be the entire height of the screen 1080
## X,Y,L,H. X&Y are top left corner position. L&W of the GUI window
if b != 0:                                                                  # if b is not equal to 0 (Y in the 0 location)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 960, 0, 960, 1080, 0)    # set to this location; X=960, Y=0, L=960, H=1080 
    print('Kinesis is not right justified... from the if statement.')
else:                                                                       # else, right justify anyways
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 960, 0, 960, 1080, 0)    # Y may be at 0, but some of the other coordinates might not be.
    print('Kinesis seems to be right justified... from the else statement.')

kin_main()
time.sleep(1)
###########################################      Assign Screenshots to Variables      ################################################
path = r"C:\Users\shaynes\OneDrive - RxSight, Inc\Desktop\OCT XY-Stage\ThorLabs Kinesis/"
k_allchecked    = "Kin-AllChecked.png"
k_allunchecked  = "Kin-AllUnChecked.png"
k_arrow         = "Kin-Arrow.png"
k_arrowns       = "Kin-ArrowNS.png"
k_check         = "Kin-Chk.png"
k_conn          = "Kin-Conn.png"
k_connn         = "Kin-Connn.png"
k_connxy        = "Kin-ConnXY.png"
k_connyx        = "Kin-ConnYX.png"
k_drag          = "Kin-Drag.png"
k_error         = "Kin-Error45318324.png"
k_home          = "Kin-Home.png"
k_home_cls      = "Kin-HomeClose.png"
k_home_dpdn     = "Kin-HomeDpDn.png"
k_not_homed     = "Kin-HomeNot.png"
k_notallconn    = "Kin-NoAllConnected.png"
k_nodevices     = "Kin-NoDevices.png"
k_nousb         = "Kin-NoUSB.png"
k_noxconn       = "Kin-NoXConn.png"
k_noxnoyconn    = "Kin-NoXNoYConn.png"
k_noyconn       = "Kin-NoYConn.png"
k_resume        = "Kin-Resume.png"
k_run           = "Kin-Run.png"
k_seqopt        = "Kin-SeqOpt.png"
k_testseq       = "Kin-TestSeq.png"
k_tseq_dpdn     = "Kin-TestSeqDpDn.png"
k_tseq_cls      = "Kin-TSeqClose.png"
k_x_sn          = "Kin-XandSN.png"
k_xdis          = "Kin-XDis.png"
k_xen           = "Kin-XEn.png"
k_xzero         = "Kin-XHome.png"
k_xsn           = "Kin-XSN.png"
k_xstart        = "Kin-Xstart.png"
k_y_sn          = "Kin-YandSN.png"
k_ydis          = "Kin-YDis.png"
k_yen           = "Kin-YEn.png"
k_yzero         = "Kin-YHome.png"
k_ysn           = "Kin-YSN.png"
k_ystart        = "Kin-Ystart.png"
k_thorlabs      = "Thorlabs.png"

## Drag log screen down (to be able to enable X when log is full, otherwise X-axis is half covered)
pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_thorlabs}'))
print('Found Thorlabs photo.')
pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_drag}'))
print('Found Drag photo')
pyautogui.moveRel(xOffset=0, yOffset=17)
print('Move down 10')
pyautogui.dragRel(xOffset=0, yOffset=200, button='left')
print('Drag it down 150')

## Check and Connect X&Y. 
while(True):                                                                                # Loop as long as this is false, can't find the no devices image
    try:                                                                                    # try and locate "Move devices here to access full functionality" on the screen. Neither X nor Y are connected.
        anodev, bnodev = pyautogui.locateCenterOnScreen(path + f'{k_nodevices}')
    except TypeError:                                                                       
        print("Something must be connected..png")
    else:                                                                                   # else gets executed if it found the try statement.
        print("No devices connected. Let's connect X&Y'")
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_conn}'))                 # press connect button

        while(True):                                                                        # loop until condition in false.
            try:                                                                            # Try and find image of the check boxes.
                anoxnoy, bnoxnoy = pyautogui.locateCenterOnScreen(path + f'{k_noxnoyconn}') # Image of X&Y axis unchecked.
            except TypeError:                                                               # execute until the check boxes popup.
                print("Could not locate the image Kin-NoXNoyConn, so let's click connect.")
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_conn}'))         # connect button
                time.sleep(2)

                while(True):                                                                # loop until condition is false.
                    try:                                                                    # try and find image of all the boxes checked.
                        aallcheck, ballcheck=pyautogui.locateCenterOnScreen(path + f'{k_allchecked}') # Image of All checked
                    except TypeError:                                                       # Execute until boxes are checked.
                        print("Could not locate the image Kin-AllChecked.png, so let's click again.")
                        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_check}'))# check top box
                        time.sleep(2) 

                    while(True):                                                            # loop until condition in false, until it can't find the image of YSN
                        try:                                                                # Try and find image of the serial number.
                            aysn, bysn=pyautogui.locateCenterOnScreen(path + f'{k_ysn}')    # Loop until the connected button has been clicked.
                        except:                                                             # Exception executed, click connect until the serial number pops up.
                            print("Could not locate the image of Kin-YSN, so let's click connect.")
                            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_connn}'))# connect
                            time.sleep(3) 
                        break
                    print("X&Y should be connected")
                    break
            break
    break
time.sleep(2)

def conn_x_or_y():
    aXconn, bXconn = pyautogui.locateCenterOnScreen(path + f'{k_xsn}')                      # write XSN X,Y coordinates of image to the variables (if it exists)
    aYconn, bYconn = pyautogui.locateCenterOnScreen(path + f'{k_ysn}')                      # write YSN X,Y coordinates of image to the variables (if it exists)
    if aXconn is True and bXconn is True:                                                   # if locating on screen returns a value, then the image is on the screen
## if XSN is true, then that means X was not connected, and from previous checks, we know Y must be connected (or USB issue)
## Check and Connect Y, we already know that 1 is connected, but both are not connected, from previous while loops.
        try:                                                                                # try and locate the image if x is connected, that would mean we need to connect Y.
            aXconn, bXconn = pyautogui.locateCenterOnScreen(path + f'{k_xsn}')              # locate SN of X.
        except TypeError:
            print("Could not locate the image Kin-XSN.png")
        else:
            print("X is connected, so lets connect Y.")
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_conn}'))             # click connect button
            time.sleep(2) 
            try:                                                                            # if the image of X and Y not connected is true, this will connect them.
                aNoYconn,bNoYconn = pyautogui.locateCenterOnScreen(path + f'{k_noyconn}')   # Only Y available to connect
                time.sleep(3)
            except TypeError:
                print("Could not locate the image Kin-NoYConn.png. Something else must the issue, Y appears to be connected")
            else:
                time.sleep(3)
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_check}'))        # click check top box
                time.sleep(3) 
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_connn}'))        # click connect button
                time.sleep(3) 
                print("Y is connected")
    elif aYconn is True and bYconn is True:                                                 # means YSN is true, so we connect YSN now  
 ## Check and Connect X
        try:                                                                                # try and locate this image on the screen of Y is connected.
            aYconn, bYconn = pyautogui.locateCenterOnScreen(path + f'{k_ysn}')
        except TypeError:
            print("Could not locate the image Kin-YSN.png")
        else:
            print("Y is connected, let's connext X'")
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_conn}'))             # click connect button
            time.sleep(2) 
            try:                                                                            # if the image of X and Y not connected is true, this will connect them.
                aNoXconn,bNoXconn = pyautogui.locateCenterOnScreen(path + f'{k_noxconn}')   # Only X not conn
                time.sleep(3)
            except TypeError:
                print("Could not locate the image Kin-NoXConn.png. Something else must the issue, X appears to be connected")
            else:
                time.sleep(3)
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_check}'))        # click the top check box
                time.sleep(3) 
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_connn}'))        # click connect button
                print("X is now connected")
                time.sleep(3) 
    else:
        print("Could not locate XSN or YSN on the screen.... Check USB and power")

## If X and Y are connected, the coordinates of the matching screenshot will be written to these assigned variables.
## Double check that X and Y are connected. No while loop b/c we don't want this to loop if its true or false.
try:                                                                                        # try and locate this image on the screen. X and Y are connected.
    aXYconn1, bXYconn1 = pyautogui.locateCenterOnScreen(path + f'{k_connxy}')
except TypeError:
    print("Could not locate the image Kin-XYConn.png, therefore X or Y is NOT connected, Let's try to connect to one.")
    conn_x_or_y()
else:
    print("XY are connected, check #1.")

try:                                                                                         # try and locate this image on the screen of Y and X are connected.
    aYXconn1, bYXconn1 = pyautogui.locateCenterOnScreen(path + f'{k_connyx}')
except TypeError:
    print("Could not locate the image Kin-YXConn.png, therefore Y or X is NOT connected. Let's try to connect to one.")
    conn_x_or_y()
else:
    print("YX are connected, check #1.")

## Double check that X and Y are connected. No while loop b/c we don't want this to loop if its true or false.
try:                                                                                        # try and locate this image on the screen. X and Y are connected.
    aXYconn2, bXYconn2 = pyautogui.locateCenterOnScreen(path + f'{k_connxy}')
except TypeError:
    print("Could not locate the image Kin-XYConn.png, therefore X & Y are NOT connected")
else:
    print("XY are connected, check #2.")
                                                                                            # Double check that Y and X are connected. No while loop b/c we don't want this to loop if its true or false.
try:                                                                                        # try and locate this image on the screen of Y and X are connected.
    aYXconn2, bYXconn2 = pyautogui.locateCenterOnScreen(path + f'{k_connyx}')
except TypeError:
    print("Could not locate the image Kin-YXConn.png, therefore Y & X are NOT connected")
else:
    print("YX are connected, check #2.")

while(True):                                                                                # loop until condition in false.
    try:                                                                                    # Try and find image of the serial number.
        k,l=pyautogui.locateCenterOnScreen(path + f'{k_ysn}')                               # Loop until the connected button has been clicked.
    except:                                                                                 # Execute clicking connect until the serial number popsup.
        print("Could not locate the image of Kin-YSN, so let's click connect.")
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_connn}'))                # connect
        time.sleep(3) 
    break
print("X&Y should be connected, check #3")

fn=''                                                                                       # needed to say " while false"
## Enable Y-axis
try:                                                                                        # Try: check if enable is on the screen, if it is then execute else
    aYen, bYen = pyautogui.locateCenterOnScreen(path + f'{k_yen}')                          # locate and write coordinates to the variables
except:                                                                                     # Except, if no enable is on the screen, then it is already enabled
    print("Couldn't find Kin-YEn image. Y is already enabled") 
else:
    pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_yen}'))                      # locate and click enable button
    print("Y is now enabled")                                                               # if it finds and clicks the Enable button on try, it will print this.

## Enable X-axis
try:                                                                                        # Try: check if enable is on the screen, if it is then execute else
    aXen, bXen = pyautogui.locateCenterOnScreen(path + f'{k_xen}')
except:                                                                                     # If no exception happened, meaning it is not enabled this block is executed.      
    print("Couldn't find Kin-XEn image. X is already enabled") 
else:
    pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_xen}'))                      # locate and click enable button
    print("X is now enabled")                                                               # if it finds the Enable button on try, it will print this.


## Double check if disable button is now visible Y-axis
try:                                                                                        # Try: check if Disable is on the screen, if it is then execute else
    aYdis,bYdis = pyautogui.locateCenterOnScreen(path + f'{k_ydis}')                        
except:                                                                                     # Except, if no Disable is on the screen, then click the center of EnY pic
    pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_yen}'))                      # locate and click enable button
    print("Couldn't find Kin-Y Disable button, need to click enable.") 
else:
    print("Found Y-Disable button. Y is enabled")                                           # if it finds the Disable button on try, it will print this.

## Double check if disable button is now visible X-axis
try:                                                                                        # Try: check if Disable is on the screen, if it is then execute else
    aXdis,bXdis = pyautogui.locateCenterOnScreen(path + f'{k_xdis}')                        
except:                                                                                     # Except, if no Disable is on the screen, then click the center of EnY pic
    pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_xen}'))                      # locate and click enable button
    print("Couldn't find Kin-X Disable button, need to click enable.") 
else:
    print("Found X-Disable button. X is enabled")                                           # if it finds the Disable button on try, it will print this.

def seq_home():                                                                                 # home function
    try:                                                                                    # try and assign the image of X is in home pos 0.000000 mm
        aXzero, bXzero = pyautogui.locateCenterOnScreen(path + f'{k_xzero}')                # assign X,Y coordinates to the image of x at 0.00000 mm
    except:                                                                                 # exception executed if image does not exist
        try:
            ahom, bhom = pyautogui.locateCenterOnScreen(path + f'{k_home}')                 # try and assign the image of home sequence already loaded
        except TypeError:                                                                   # exception executed if image does not exist 
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))            # click on "Open" drop down arrow
            time.sleep(2)                                                                   # pause
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_home_dpdn}'))        # click on home sequence from drop down
            print("Loading Home Sequence.") 
            time.sleep(2)                                                                   # pause
            pyautogui.moveRel(xOffset=-100, yOffset=140)                                    # move X,Y relative to current position
            pyautogui.click()                                                               # click "Run" button
            print("Pausing to allow to home...")          
            time.sleep(9)                                                                   # pause 8 seconds to allow to home.
            print("Closing home sequence")
            ahomcls, bhomcls = pyautogui.locateCenterOnScreen(path + f'{k_home_cls}')       # assign the image of home sequence already loaded with red x
            pyautogui.moveTo(x=ahomcls, y=bhomcls)                                          # move to image of home sequence with red x
            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
            pyautogui.click()                                                               # click "close" button
        else:                                                                               # else means try image was found (home sequence already loaded)
            print("Home sequence is already open, lets press Run.")
            pyautogui.moveTo(x=ahom, y=bhom)                                                # move to image of home sequence
            pyautogui.moveRel(xOffset=-20, yOffset=100)                                     # move X,Y relative to current position
            pyautogui.click()                                                               # click "Run" button
            print("Pausing to allow to home...")          
            time.sleep(9)                                                                   # pause 8 seconds to allow to home.
            print("Closing home sequence")
            ahomcls, bhomcls = pyautogui.locateCenterOnScreen(path + f'{k_home_cls}')       # assign the image of home sequence already loaded with red x
            pyautogui.moveTo(x=ahomcls, y=bhomcls)                                          # move to image of home sequence with red x
            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
            pyautogui.click()                                                               # click "close" button
    else:                                                                                   # else means try image was found (X is already at starting pos 0.000000 mm)
        print("X is already at 0.000000 mm")

    ## if X is already at 0.000000 mm , but Y is not, we will find out here..
    try:                                                                                    # try and assign the image of Y is in starting pos 0.000000 mm
        aYzero, bYzero = pyautogui.locateCenterOnScreen(path + f'{k_yzero}')                # if image exists, assign X,Y coordinates to the image of Y at 0.00000 mm
    except:                                                                                 # exception executed if image does not exist                                               
        try:
            ahom, bhom = pyautogui.locateCenterOnScreen(path + f'{k_home}')                 # try and assign the image of home sequence already loaded
        except TypeError:                                                                   # if no exception, means no home sequence is loaded. This gets executed if image does not exist
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))            # click on "Open" drop down arrow
            time.sleep(2)                                                                   # pause
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_home_dpdn}'))        # click on home sequence from drop down
            print("Loading Home Sequence.") 
            time.sleep(2)                                                                   # pause
            pyautogui.moveRel(xOffset=-100, yOffset=140)                                    # move X,Y relative to current position
            pyautogui.click()                                                               # click "Run" button
            print("Pausing to allow to home...")          
            time.sleep(9)                                                                   # pause 8 seconds to allow to home.
            print("Closing home sequence")
            ahomcls, bhomcls = pyautogui.locateCenterOnScreen(path + f'{k_home_cls}')       # assign the image of home sequence already loaded with red x
            pyautogui.moveTo(x=ahomcls, y=bhomcls)                                          # move to image of home sequence with red x
            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
            pyautogui.click()                                                               # click "close" button
        else:                                                                               # else means try image was found (home sequence already loaded)
            print("Home sequence is already open, lets press Run.")
            pyautogui.moveTo(x=ahom, y=bhom)                                                # move to image of home sequence
            pyautogui.moveRel(xOffset=-20, yOffset=100)                                     # move X,Y relative to current position
            pyautogui.click()                                                               # click "Run" button
            print("Pausing to allow to home...")          
            time.sleep(9)                                                                   # pause 8 seconds to allow to home.
            print("Closing home sequence")
            ahomcls, bhomcls = pyautogui.locateCenterOnScreen(path + f'{k_home_cls}')       # assign the image of home sequence already loaded with red x
            pyautogui.moveTo(x=ahomcls, y=bhomcls)                                          # move to image of home sequence with red x
            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
            pyautogui.click()                                                               # click "close" button
    else:                                                                                   # else means try image was found (Y is already at starting pos 0.000000 mm)
        print("Y is already at 0.000000 mm")
    ## upon startup, X and Y can say 0.00000 mm on the screen, but still be somewhere in space. In that case. Check for "Not Homed" image
    try:                                                                                    # try and assign the image of "Not Homed"
        anhome, bnhome = pyautogui.locateCenterOnScreen(path + f'{k_not_homed}')            # assign X,Y coordinates to the image of "Not Homed"
    except TypeError:                                                                       # exception executed if image does not exist
        print("No image of Not Homed exists")                                                              # click "close" button
    else:                                                                                   # else means not homed image was found
        print("Device is not homed, lets home it.")
        axsn, bxsn = pyautogui.locateCenterOnScreen(path + f'{k_xsn}')                      # assign the coordinates of XSN image
        aysn, bysn = pyautogui.locateCenterOnScreen(path + f'{k_ysn}')                      # assign the coordinates of YSN image
        pyautogui.moveTo(x=axsn, y=bxsn)                                                    # move to image of xsn
        pyautogui.moveRel(xOffset=-170, yOffset=110)                                        # move X,Y relative to current position
        pyautogui.click()                                                                   # click "start" button
        time.sleep(2)                                                                       # pause
        pyautogui.moveTo(x=aysn, y=bysn)                                                    # move to image of ysn
        pyautogui.moveRel(xOffset=-170, yOffset=110)                                        # move X,Y relative to current position
        pyautogui.click()                                                                   # click "start" button
        time.sleep(75)                                                                      # pause for home time 1 minute 12 seconds
        print("Should be homed now.")

# shouldn't ever need to do this. Using test sequence it should go to start, beginning position everytime.
#def seq_start():                                                                            # start function for Lumedica loops to not have to return to home position
#    try:                                                                                    # try and assign the image of X is in starting pos 132.700000 mm
#        asXstart, bsXstart = pyautogui.locateCenterOnScreen(path + f'{k_xstart}')           # assign X,Y coordinates to the image of x at 132.70000 mm
#    except:                                                                                 # exception executed if image does not exist
#        try:                                                                                # try and assign the image of start sequence already loaded
#            aseqstart, bseqstart = pyautogui.locateCenterOnScreen(path + f'{k_start}')              
#        except TypeError:                                                                   # exception executed if image does not exist 
#            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))            # click on "Open" drop down arrow
#            time.sleep(2)                                                                   # pause
#            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_start_dpdn}'))       # click on start sequence from drop down
#            print("Loading Start Sequence.") 
#            time.sleep(2)                                                                   # pause
#            pyautogui.moveRel(xOffset=-100, yOffset=140)                                    # move X,Y relative to current position
#            pyautogui.click()                                                               # click "Run" button
#            print("Pausing to allow to start...")
#            time.sleep(9)                                                                   # pause 8 seconds to allow to move to start.
#            print("Closing start sequence")
#            aseqstarcls, bseqstarcls = pyautogui.locateCenterOnScreen(path + f'{k_start_cls}')# assign the image of start sequence already loaded with red x
#            pyautogui.moveTo(x=aseqstarcls, y=bseqstarcls)                                  # move to image of start sequence with red x
#            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
#            pyautogui.click()                                                               # click "close" button
#        else:                                                                               # else means try image was found start sequence already loaded)
#            pyautogui.moveTo(x=astart, y=bstart)                                            # move to image of start sequence
#            pyautogui.moveRel(xOffset=-100, yOffset=140)                                    # move X,Y relative to current position
#            pyautogui.click()                                                               # click "Run" button
#            print("Pausing to allow to start...")
#            time.sleep(9)                                                                   # pause 8 seconds to allow to move to start.
#            print("Closing start sequence")
#            astarcls, bstarcls = pyautogui.locateCenterOnScreen(path + f'{k_start_cls}')    # assign the image of start sequence already loaded with red x
#            pyautogui.moveTo(x=astarcls, y=starcls)                                         # move to image of start sequence with red x
#            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
#            pyautogui.click()   
#    else:                                                                                   # else means try image was found (X is already at starting pos 0.000000 mm)
#        print("X is already at 132.700000 mm")

#    ## if X is already at 132.700000 mm , but Y is not at 25.7, we will find out here..
#    try:                                                                                    # try and assign the image of Y is in starting pos 25.700000 mm
#        asYstart, bsYstart = pyautogui.locateCenterOnScreen(path + f'{k_ystart}')             # assign X,Y coordinates to the image of x at 25.70000 mm
#    except:                                                                                 # exception executed if image does not exist
#        try:                                                                                # try and assign the image of start sequence already loaded
#            astart, bstart = pyautogui.locateCenterOnScreen(path + f'{k_start}')              
#        except TypeError:                                                                   # exception executed if image does not exist 
#            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))            # click on "Open" drop down arrow
#            time.sleep(2)                                                                   # pause
#            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_start_dpdn}'))       # click on start sequence from drop down
#            print("Loading Start Sequence.") 
#            time.sleep(2)                                                                   # pause
#            pyautogui.moveRel(xOffset=-100, yOffset=140)                                    # move X,Y relative to current position
#            pyautogui.click()                                                               # click "Run" button
#            print("Pausing to allow to start...")
#            time.sleep(9)                                                                   # pause 8 seconds to allow to move to start.
#            print("Closing start sequence")
#            astarcls, bstarcls = pyautogui.locateCenterOnScreen(path + f'{k_start_cls}')    # assign the image of start sequence already loaded with red x
#            pyautogui.moveTo(x=astarcls, y=starcls)                                         # move to image of start sequence with red x
#            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
#            pyautogui.click()                                                               # click "close" button
#        else:                                                                               # else means try image was found start sequence already loaded)
#            pyautogui.moveTo(x=astart, y=bstart)                                            # move to image of start sequence
#            pyautogui.moveRel(xOffset=-100, yOffset=140)                                    # move X,Y relative to current position
#            pyautogui.click()                                                               # click "Run" button
#            print("Pausing to allow to start...")
#            time.sleep(9)                                                                   # pause 8 seconds to allow to move to start.
#            print("Closing start sequence")
#            astarcls, bstarcls = pyautogui.locateCenterOnScreen(path + f'{k_start_cls}')    # assign the image of start sequence already loaded with red x
#            pyautogui.moveTo(x=astarcls, y=starcls)                                         # move to image of start sequence with red x
#            pyautogui.moveRel(xOffset=90, yOffset=0)                                        # move X,Y relative to current position
#            pyautogui.click()   
#    else:                                                                                   # else means try image was found (X is already at starting pos 0.000000 mm)
#        print("Y is already at 25.700000 mm")

def seq_test():
    try:                                                                                        # try and assign the image of X is in starting pos 132.700000 mm
        atXstart, btXstart = pyautogui.locateCenterOnScreen(path + f'{k_xstart}')               # if image exists, assign X,Y coordinates to the image of X at 132.70000 mm
    except:                                                                                     # exception executed if image does not exist
        try: 
            atxseq, btxseq = pyautogui.locateCenterOnScreen(path + f'{k_testseq}')              # try and assign the image of test sequence already loaded
        except TypeError:                                                                       # if no exception, means no test sequence is loaded. This gets executed if image does not exist
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))                # click on "Open" drop down arrow
            time.sleep(2)                                                                       # pause
            pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_tseq_dpdn}'))            # click on test sequence from drop down
            time.sleep(2)                                                                       # pause
            pyautogui.moveRel(xOffset=-100, yOffset=120)                                        # move X,Y relative to current position
            pyautogui.click()                                                                   # click "Run" button
            time.sleep(7)                                                                       # pause at least 5-7 seconds to  move to start
            print("Loading Test Sequence.")  
        else:                                                                                   # else means try image was found (test sequence already loaded)
            pyautogui.moveTo(x=atxseq, y=btxseq)                                                # move to test sequence image
            pyautogui.moveRel(xOffset=-72, yOffset=120)                                         # move X,Y relative to current position
            pyautogui.click()                                                                   # click "Run" button
            time.sleep(7)                                                                       # pause at least 5-7 seconds to move to start
            print("Test sequence is already open, lets press Run.")
    else:                                                                                       # else means try image was found (X is already at starting pos 132.700000 mm)
        print("X-axis already in starting position. Now let's check Y.")
        try:                                                                                        # try and assign the image of Y in starting pos 25.700000 mm
            atYstart, btYstart = pyautogui.locateCenterOnScreen(path + f'{k_ystart}')               # if image exists, assign X,Y coordinates to the image of Y at 25.70000 mm
        except:                                                                                     # exception executed if image does not exist
            try: 
                atyseq, btyseq = pyautogui.locateCenterOnScreen(path + f'{k_testseq}')              # try and assign the image of test sequence already loaded
            except TypeError:                                                                       # if no exception, means no test sequence is loaded. This gets executed if image does not exist
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))                # click on "Open" drop down arrow
                time.sleep(2)                                                                       # pause
                pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_tseq_dpdn}'))            # click on test sequence from drop down
                time.sleep(2)                                                                       # pause
                pyautogui.moveRel(xOffset=-100, yOffset=120)                                        # move X,Y relative to current position
                pyautogui.click()                                                                   # click "Run" button
                time.sleep(7)                                                                       # pause at least 5-7 seconds to move to start
                print("Loading Test Sequence.")  
            else:                                                                                   # else means try image was found (test sequence already loaded)
                pyautogui.moveTo(x=atseq, y=btseq)                                                  # move to test sequence image
                pyautogui.moveRel(xOffset=-72, yOffset=120)                                         # move X,Y relative to current position
                pyautogui.click()                                                                   # click "Start" button
                time.sleep(7)                                                                       # pause at least 5-7 seconds to move to start
                print("Test sequence is already open, lets press start.")
        else:                                                                                       # else means try image was found (X is already at starting pos 132.700000 mm)
            print("Y-axis already in starting position.")
        print("Checked, X & Y were already in starting position, so test did not get started.")
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_arrow}'))                # click on "Open" drop down arrow
        time.sleep(2)                                                                       # pause
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_tseq_dpdn}'))            # click on test sequence from drop down
        time.sleep(2)                                                                       # pause
        pyautogui.moveRel(xOffset=-100, yOffset=120)                                        # move X,Y relative to current position
        pyautogui.click()                                                                   # click "Run" button
        time.sleep(7)                                                                       # pause at least 5-7 seconds to move to start
        print("Loading Test Sequence.")  

atseqcls, btseqcls = pyautogui.locateCenterOnScreen(path + f'{k_tseq_cls}')                 # assign X,Y coordinates to the image of test sequence with red x for closing after test finished.

# If there are less than 25 samples in a set, we need a way to cancel the test sequence and restart the test for the next .
def cancel_test():
    kin_pop()                                                                   # bring popup to foreground
    try: 
        acancel, bcancel = pyautogui.locateCenterOnScreen(path + f'{k_cancel}') # try and find the resume button
    except:                                                                     # except executed if no image was found
        print("No resume button found")
    else:                                                                       # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_cancel}'))
        print("Should be resuming kinesis to move to next sample")

############################################     DETERMINING # OF LOOPS        ##################################################  
fill_arr = 30       # TEST VALUE
samp_arr_raw = []      # TEST VALUE
for fa in range(fill_arr):      # TEST VALUE
    samp_arr_raw.append(fa)      # TEST VALUE

print("Sample Array Raw is: ", samp_arr_raw)

samp_len_25 = 0                                                 # if this is 0 & loop does not get updated by if statement, loop will not execute.
samp_len_50 = 0                                                 # if this is 0 & loop does not get updated by if statement, loop will not execute.
samp_len_75 = 0                                                 # if this is 0 & loop does not get updated by if statement, loop will not execute.
samp_len_100 = 0                                                # if this is 0 & loop does not get updated by if statement, loop will not execute.
samp_len_125 = 0                                                # if this is 0 & loop does not get updated by if statement, loop will not execute.
samp_len = len(samp_arr_raw)
print("Sample Length is: ", samp_len)

if samp_len <= 25:                                                                      # if the array length is less than 25
    print("Sample array length is less than 25 samples.")
    samp_len_25 = samp_len                                                              # Set the numeric sample length to variable samp_len_25 to use for the loop
    print("samp_len_25 is: ", samp_len_25, ". Therefore 0 outer loops.")
elif samp_len > 25 and samp_len <= 50:                                                  # if the array length is between 26 and 50, prob most common
    print("Sample array length is between 26 and 50 samples.")
    samp_len_25 = 25                                                                    # first inner loop
    samp_len_50 = samp_len - 25                                                         # second inner loop
    print("samp_len_25 is: ", samp_len_25, ", samp_len_50 is: ", samp_len_50, ". Therefore 1 outer loop.")
elif samp_len > 50 and samp_len <= 75:                                                  # if the array length is between 51 and 75
    print("Sample array length is between 51 and 75 samples.")
    samp_len_25 = 25                                                                    # first inner loop
    samp_len_50 = 25                                                                    # second inner loop
    samp_len_75 = samp_len - 50                                                         # third inner loop
    print("samp_len_25 is: ", samp_len_25, ", samp_len_50 is: ", samp_len_50)
    print("samp_len_75 is: ", samp_len_75, ". Therefore 2 outer loops.")
elif samp_len > 75 and samp_len <= 100:                                                 # this is for scalability. If the array length is between 76 and 100
    print("Sample array length is between 76 and 100 samples.")
    samp_len_25 = 25                                                                    # first inner loop
    samp_len_50 = 25                                                                    # second inner loop
    samp_len_75 = 25                                                                    # third inner loop
    samp_len_100 = samp_len - 75                                                        # fourth inner loop
    print("samp_len_25 is: ", samp_len_25, ", samp_len_50 is: ", samp_len_50)
    print("samp_len_75 is: ", samp_len_75, ", samp_len_100 is: ", samp_len_100)
    print("Therefore 3 outer loops.")
elif samp_len > 101 and samp_len <= 125:                                                # this is for scalability. If the array length is between 101 and 125
    print("Sample array length is between 101 and 125 samples.")
    samp_len_25 = 25                                                                    # first inner loop
    samp_len_50 = 25                                                                    # second inner loop
    samp_len_75 = 25                                                                    # third inner loop
    samp_len_100 = 25                                                                   # fourth inner loop
    samp_len_125 = samp_len - 100                                                       # fifth inner loop
    print("samp_len_25 is: ", samp_len_25, ", samp_len_50 is: ", samp_len_50)
    print("samp_len_75 is: ", samp_len_75, ", samp_len_100 is: ", samp_len_100)
    print("samp_len_125 is: ", samp_len_125, ", Therefore 4 outer loops.")
else:
    print("Too many samples, software not configured for this.")

############################################        START OUTER SET LOOP          ##################################################  
index_015 = 0                                       # count is for the sample index, naming each sample based on the array. Needs to be outside the loop
index_040 = 0                                       # count is for the sample index, naming each sample based on the array. Needs to be outside the loop

seq_test()                                                                              # start test sequence
for len25_015 in range(samp_len_25):                                                    # will loop from 0 to samp_len_25
    print("Testing the 015 for loop of len25.", len25_015)
    print("Index 015 test: ", samp_arr_raw[index_015])
    ############################################          START  015   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
    #lum_mini()         # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")
    index_015 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len25_040 in range(samp_len_25):                                                    # will loop from 0 to samp_len_25
    print("Testing the 040 for loop of len25.", len25_040)
    print("Index 040 test: ", samp_arr_raw[index_040])
    index_040 +=1
    ############################################          START  040   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
    #lum_mini()         # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")
    index_040 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len50_015 in range(samp_len_50):                                                    # will loop from 0 to samp_len_50
    print("Testing the 015 for loop of len50.", len50_015)
    print("Index 015 test: ", samp_arr_raw[index_015])
    ############################################          START  015   TEST        ##################################################  
                            ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)           # this pause button simulates Lumedica performing measurements.
    #lum_mini()             # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")
    index_015 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len50_040 in range(samp_len_50):                                                    # will loop from 0 to samp_len_50
    print("Testing the 040 for loop of len50.", len50_040)
    print("Index 040 test: ", samp_arr_raw[index_040])
    ############################################          START  040   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
                        # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")    
    index_040 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len75_015 in range(samp_len_75):                                                    # will loop from 0 to samp_len_75
    print("Testing the 015 for loop of len75.", len75_015)
    print("Index 015 test: ", samp_arr_raw[index_015])
    ############################################          START  015   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
    #lum_mini()         # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")
    index_015 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len75_040 in range(samp_len_75):                                                      # will loop from 0 to samp_len_75
    print("Testing the 040 for loop of len75.", len75_040)
    print("Index 040 test: ", samp_arr_raw[index_040])
    ############################################          START  040   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
                        # minimize Lumedica
    kin_pop()                                                                       # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')           # try and find the resume button
    except:                                                                         # except executed if no image was found
        print("No resume button found")
    else:                                                                           # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")    
    index_040 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len100_015 in range(samp_len_100):                                                    # will loop from 0 to samp_len_100
    print("Testing the 015 for loop of len100.", len100_015)
    print("Index 015 test: ", samp_arr_raw[index_015])
    ############################################          START  015   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
    #lum_mini()         # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")
    index_015 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len100_040 in range(samp_len_100):                                                    # will loop from 0 to samp_len_100
    print("Testing the 040 for loop of len100.", len100_040)
    print("Index 040 test: ", samp_arr_raw[index_040])
    ############################################          START  040   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
                        # minimize Lumedica
    kin_pop()                                                                       # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')           # try and find the resume button
    except:                                                                         # except executed if no image was found
        print("No resume button found")
    else:                                                                           # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")    
    index_040 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len125_015 in range(samp_len_125):                                                  # will loop from 0 to samp_len_125
    print("Testing the 015 for loop of len125.", len125_015)
    print("Index 015 test: ", samp_arr_raw[index_015])
    ############################################          START  015   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
    #lum_mini()         # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")
    index_015 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.
seq_test()                                                                              # start test sequence

for len125_040 in range(samp_len_125):                                                  # will loop from 0 to samp_len_125
    print("Testing the 040 for loop of len125.", len125_040)
    print("Index 040 test: ", samp_arr_raw[index_040])
    ############################################          START  040   TEST        ##################################################  
                        ## Do Lumedica measurements here, or call the function that does them
    time.sleep(2)       # this pause button simulates Lumedica performing measurements.
                        # minimize Lumedica
    kin_pop()                                                                           # bring popup to foreground
    try: 
        ares, bres = pyautogui.locateCenterOnScreen(path + f'{k_resume}')               # try and find the resume button
    except:                                                                             # except executed if no image was found
        print("No resume button found")
    else:                                                                               # else executed if image was found
        pyautogui.click(pyautogui.locateCenterOnScreen(path + f'{k_resume}'))
        print("Should be resuming kinesis to move to next sample")    
    index_040 +=1

cancel_test()                                                                           # need this incase there are less than 25 samples.

atseq, btseq = pyautogui.locateCenterOnScreen(path + f'{k_testseq}')                    # close test sequence
pyautogui.moveTo(x=atseq, y=btseq) 
pyautogui.moveRel(xOffset=90, yOffset=-10)                                              # Move to "x" button and close home sequence.
pyautogui.click()

home()  # Return XY stage to home before closing.
time.sleep(10)                                                                          # wait 10 seconds to home

os.system("TaskKill /F /IM Thorlabs.MotionControl.Kinesis.exe")                        # close Kinesis
exit()
