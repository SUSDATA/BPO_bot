import os
from dotenv import load_dotenv
import subprocess
import pyperclip
import pyautogui
from genericfunctions import (pressingKey,make_window_visible)
from time import sleep

# Open VPN Client and Connect to NAE VPN
def openAndConnectVPNForticlient():    
    forticlient_connected = None
    subprocess.Popen('C:\\Program Files\\Fortinet\\FortiClient\\FortiClient.exe')
    sleep(3)
    pressingKey('tab',2)
    sleep(2)
    pyautogui.write('dgard')    
    sleep(1)
    pressingKey('tab',1)
    sleep(1)
    pyautogui.write('Colombia2023*')
    pressingKey('enter')
    while forticlient_connected is None:
        forticlient_connected = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/forti_client_connected.png', grayscale = True,confidence=0.9)    
    print("FortiClient Connected!")
    sleep(1)
    pyautogui.getWindowsWithTitle("FortiClient")[0].minimize()

# Open CRM Onyx App
def openCRM():    
    crm_login_app = crm_login_app_2 = None
    crmAttempts = 0

    sleep(1)
    pressingKey('win')
    sleep(1)
    pyautogui.write('Nuevo CRM')
    sleep(1)
    pressingKey('enter')
   
    # while crm_login is None:
    #     crm_login = pyautogui.locateOnScreen('C:/RPA BPO Automation/Version 1.0/assets/crm_login_window.png', grayscale = True,confidence=0.9)    
    # print("CRM Login GUI is present!")

    #pyautogui.click(906,743)
    while crm_login_app is None and crm_login_app_2 is None:        
        crm_login_app = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_login_window.png', grayscale = True,confidence=0.9)
        crm_login_app_2 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_login_window_v2.png', grayscale = True,confidence=0.9)
        sleep(0.6)
        print("buscando crm login window")
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Control de Acceso Unificado CRM")[0].minimize()
            pyautogui.getWindowsWithTitle("Control de Acceso Unificado CRM")[0].maximize()
            crmAttempts = 0
        crmAttempts = crmAttempts + 1

    print("CRM dashboard login is present!")    
# Connect CRM Onyx App    
def connectToCRM():    
    crm_select_app = None
    crm_dashboard = None       
    
    pressingKey('tab',2)
    sleep(1)
    pyautogui.write(os.getenv('USER_CRM'))
    sleep(1)
    pressingKey('tab') 
    pyperclip.copy(os.getenv('PASS_CRM'))
    sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    sleep(1)
    pressingKey('tab')
    sleep(2)
    pressingKey('enter')
    
    while crm_select_app is None:        
        crm_select_app = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/select_app.png', grayscale = True,confidence=0.9)
    print("CRM select box is present!")    
    pressingKey('enter')
    pressingKey('tab',4)
    pressingKey('enter')
    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)    
    print("CRM Dashboard GUI is present!")    
    