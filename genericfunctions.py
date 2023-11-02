import pyautogui
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
username = get_username_os()

def pressingKey(key,times = 1):
    for i in range(times):
        sleep(0.6)
        pyautogui.press(key)     

def showDesktop():
    pyautogui.keyDown('win')
    pyautogui.keyDown('d')
    pyautogui.keyUp('win')
    pyautogui.keyUp('d')

def selectToEnd(times):
    for i in range(times):
        pyautogui.keyDown('shift')
        pyautogui.keyDown('fn')
        pyautogui.keyDown('end')
        pyautogui.keyUp('shift')
        pyautogui.keyUp('fn')
        pyautogui.keyUp('end')
        
def make_noise():
    duration = 1000  # milliseconds
    freq = 1140  # Hz
    winsound.Beep(freq, duration)
    winsound.Beep(freq, duration)
    winsound.Beep(freq, duration)


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
