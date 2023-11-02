import subprocess
import pyautogui
from genericfunctions import (showDesktop,pressingKey)
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
        forticlient_connected = pyautogui.locateOnScreen('C:/RPA BPO Automation/Version 1.0/assets/forti_client_connected.png', grayscale = True,confidence=0.9)    
    print("FortiClient Connected!")
    sleep(1)
    pyautogui.getWindowsWithTitle("FortiClient")[0].minimize()

# Open CRM Onyx App
def openCRM():
    crm_onyx_icon = None
    winKeyAttempts = 0
    #subprocess.Popen('C:\\Users\\dgard\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Global Hitss\\Global Hitss\\Nuevo CRM')        
    sleep(1)
    pressingKey('win')
    sleep(1)
    pyautogui.write('Nuevo CRM')    
    sleep(1)
    pressingKey('enter')
    # while crm_onyx_icon is None:
    #     crm_onyx_icon = pyautogui.locateOnScreen('C:/RPA BPO Automation/Version 1.0/assets/WIN_CRM_ONYX_SNAPSHOT.png', grayscale = True,confidence=0.9)
    #     sleep(1)
    #     if winKeyAttempts == 8:
    #         pressingKey('win')
    #         winKeyAttempts = 0
    #     winKeyAttempts = winKeyAttempts + 1
    #     print(winKeyAttempts)
            
    # sleep(1)    
    # crm_onyx_icon_punto_X,crm_onyx_icon_punto_y = pyautogui.center(crm_onyx_icon)
    # pyautogui.click(crm_onyx_icon_punto_X, crm_onyx_icon_punto_y)        
    
# Connect CRM Onyx App    
def connectToCRM():
    crm_login = None
    crm_select_app = None
    crm_dashboard = None
    
    while crm_login is None:
        crm_login = pyautogui.locateOnScreen('C:/RPA BPO Automation/Version 1.0/assets/crm_login_window.png', grayscale = True,confidence=0.9)    
    print("CRM Login GUI is present!")
    
    pressingKey('tab',2)
    sleep(1)
    pyautogui.write('EC5708D')
    sleep(1)
    pressingKey('tab')
    pyautogui.write('Django25//')
    sleep(1)
    pressingKey('tab')
    sleep(2)
    pressingKey('enter')
    
    while crm_select_app is None:        
        crm_select_app = pyautogui.locateOnScreen('C:/RPA BPO Automation/Version 1.0/assets/select_app.png', grayscale = True,confidence=0.9)
    print("CRM select box is present!")
    
    pressingKey('enter')
    pressingKey('tab',4)
    pressingKey('enter')
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/RPA BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)    
    print("CRM Dashboard GUI is present!")
    print("CRM open and connected! You can start ")
    
