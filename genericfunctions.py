import pyautogui
import win32gui
import win32com.client
import winsound
import os
from time import sleep
from datetime import datetime

# GENERIC FUNCTIONS

def getCurrentDateAndTime():
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    return current_time    

def get_username_os():
   return os.getenv("USERNAME")

def pressingKey(key,times = 1):
    for i in range(times):
        sleep(0.6)
        pyautogui.press(key)     

def showDesktop():
    pyautogui.keyDown('win')
    pyautogui.keyDown('d')
    pyautogui.keyUp('win')
    pyautogui.keyUp('d')

def selectToEnd():    
    pyautogui.keyDown('shift')
    sleep(0.1)
    pyautogui.press('end')              
    sleep(0.1)
    pyautogui.keyUp('shift')
        
def make_noise():
    duration = 1000  # milliseconds
    freq = 1140  # Hz
    winsound.Beep(freq, duration)
    winsound.Beep(freq, duration)
    winsound.Beep(freq, duration)

def make_window_visible(target_window):
    sleep(1)
    count = 1
    encontrado = None
    cantVentanas = len(pyautogui.getAllWindows())    
    
    while encontrado is None:
        window_title = pyautogui.getActiveWindowTitle()
        if target_window not in window_title:
            count += 1
            with pyautogui.hold('alt'):
                pyautogui.press('tab')
                for _ in range(0, count):                    
                    pyautogui.press('left')
        else:
            print("La ventana " + target_window + " es ahora visible y esta activa!")
            encontrado = window_title
        sleep(0.25)

#List all the active windows
def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
            print ( hex( hwnd ), win32gui.GetWindowText( hwnd ) )
    win32gui.EnumWindows( winEnumHandler, None )
    return 0

#List all the active windows
def list_all_active_windows():
    for x in pyautogui.getAllWindows():  
        print(x.title)
    while True:
        sleep(1)
        print(pyautogui.getActiveWindowTitle())

#***************************************#
#************* SEND EMAIL  *************#
#***************************************#

def sendEmail(subject,to,cc,body,attachedFile,attachedFile2):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= subject
    newmail.To=to
    newmail.CC=cc 
    newmail.Body= body 
    if attachedFile != 'null':        
        newmail.Attachments.Add(attachedFile)     
    if attachedFile2 != 'null':
        newmail.Attachments.Add(attachedFile2)
    
    # To display the mail before sending it
    # newmail.Display() 
    newmail.Send()
