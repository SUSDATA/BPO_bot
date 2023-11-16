import pyautogui
from time import sleep
from genericfunctions import (pressingKey,selectToEnd,terminateProcess)

def formatDate(date):      
    return date.strftime("%d/%m/%Y")

def openAndSearchOnCRM(incidentId):
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = None
    crmAttempts = 0
       
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)            
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
        print(crmAttempts)
        
    print("CRM Dashboard GUI is present!")
    sleep(1)
    pressingKey('f2')
    sleep(1)
    pyautogui.write(incidentId)
    pressingKey('enter')
    sleep(1)

    # Validate edit_incident view is visible and on focus    
    while crm_edit_incident is None:
        crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
    print("CRM Edit Incident button is present!")
    crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
    pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
    
    # Validate assign user pop up is visible and on focus   
    while crm_assign_user is None:
        crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
    print("CRM Confirm Assign Pop Up is present!")        
    sleep(1)
    pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
    sleep(1)
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    sleep(1)
    
    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)
    # sleep(0.5)
    # pyautogui.click(crm_save_incident_x, crm_save_incident_y)
    
    sleep(1)
    
     # Validate assign user pop up is visible and on focus 
    crm_assign_user = None  # reset variable
    while crm_assign_user is None:
        crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
    print("CRM Confirm Assign Pop Up is present!")        
    sleep(1)
    pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
    sleep(1)
    
    # Validate save OTP successfully  pop up is visible and on focus   
    while crm_otp_saved_sucessfully is None:
        crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
    print("CRM OTP Saved Succesfully Pop Up is present!")
    sleep(1)
    pressingKey('enter')
    sleep(1)
    
    # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
    # Close Edit Incident View 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
    sleep(1.5)    

def searchAndUpdateUser(incidentId,ususarioAsignado):    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = crm_warning_message = None
    asignado_a = mod_consulta_popup = None 
    crmAttempts = 0   
        
    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)            
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
        print(crmAttempts)
        
    print("CRM Dashboard GUI is present!")    
    sleep(1)    
    pressingKey('f2')
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
    print("mod_consulta_popup field is present and detected on GUI screen!")
    pyautogui.write(incidentId)
    sleep(1)
    pressingKey('enter')    
    
    #/////////////////////////////////// EDIT VIEW CRM /////////////////////////////////////////// 
    
    # Validate edit_incident view is visible and on focus    
    while crm_edit_incident is None:        
        crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
    print("CRM Edit Incident button is present!")
    crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
    pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
    
    # Validate assign user pop up is visible and on focus   
    while crm_assign_user is None:        
        crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
    print("CRM Confirm Assign Pop Up is present!")        
    sleep(1)
    pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
    sleep(1)
    
    # Maximize CRM Edit window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    sleep(1)
    
    #/////////////////////////////////// UPDATE USER ASSIGNED /////////////////////////////////////  
    # Identify which UI is present on screen. Then focus on USER ASSIGNATION FIELD and Paste the name 
    # of the user passed as parameters
            
    while asignado_a is None:        
        asignado_a = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignado_a.png', grayscale = True,confidence=0.85)    
    print("asignado_a field is present and detected on GUI screen!")
    pyautogui.click(x=301, y=263)
    pyautogui.write(ususarioAsignado)
    pressingKey('tab')
    #/////////////////////////////////// CLOSING CRM EDIT/DETAILS VIEW ///////////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)
    sleep(1)        
    
    while crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_warning_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:            
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
        
        # Validate assign user pop up is visible and on focus 
        crm_assign_user = None  # reset variable
        while crm_assign_user is None:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
        print("CRM Confirm Assign Pop Up is present!")        
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
        sleep(1)
        
         # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
        # Close Edit Incident View 
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        return 0
    else:
        sleep(1)
        pressingKey('tab')
        sleep(1)
        pressingKey('enter')
        print("CRM empty_warning_message_crm pop up is present!")
        
        # Validate assign user pop up is visible and on focus 
        crm_assign_user = None  # reset variable
        while crm_assign_user is None:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
        print("CRM Confirm Assign Pop Up is present!")        
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
        sleep(1)
        
        # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        
        # You must Validate assign user pop up is visible and on focus
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs         
        print("n key has been pressed!")  
        return 0
                        
def searchAndUpdateDates(incidentId,fechaProgramacion,fechaCompromiso):
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = crm_warning_message = crm_ot_blocked_message = None     
    vista_detalle_fecha_v1 = vista_detalle_fecha_v2 = vista_detalle_fecha_v3 = vista_detalle_fecha_v4 = vista_detalle_fecha_v5 = vista_detalle_fecha_v6 = mod_consulta_popup = None 
    crmAttempts = 0
    
    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
        print(crmAttempts)
        
    print("CRM Dashboard GUI is present!")    
    sleep(1)
    pressingKey('f2')

    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(1)
    pressingKey('enter')
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:            
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        print("CRM crm_ot_blocked_message or crm_warning_message pop up are not present!")                   
        
        # Validate edit_incident view is visible and on focus            
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("CRM Edit Incident button is present!")
        crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
        pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
    
        # Validate assign user pop up is visible and on focus
        # crm_assign_user = None  # reset variable   
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)
           
    else:
        if crm_warning_message is None:
            pressingKey('enter')
        else:  
            pressingKey('enter')
            sleep(2)
            pressingKey('enter')  

        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("CRM Edit Incident button is present!")
        crm_edit_incident = None #reset variable
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()
        return 1    
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    sleep(1)
    print("CRM Edit view Window was maximized!")    
    #/////////////////////////////////// UPDATE DATES /////////////////////////////////////  
    # Identify which UI is present on screen. Then focus on dates text input fields depending on the UI Variation and Paste the dates passed as parameters on the corresponding fields    
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(19,495)
    sleep(1)
    
    crmAttempts = 0
    while vista_detalle_fecha_v1 is None and vista_detalle_fecha_v2 is None and vista_detalle_fecha_v3 is None and vista_detalle_fecha_v4 is None and vista_detalle_fecha_v5 is None and vista_detalle_fecha_v6 is None:
        vista_detalle_fecha_v1 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/vista_detalles_fecha_V1.png', grayscale = True,confidence=0.9)
        vista_detalle_fecha_v2 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/vista_detalles_fecha_V2.png', grayscale = True,confidence=0.9)
        vista_detalle_fecha_v3 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/vista_detalles_fecha_V3.png', grayscale = True,confidence=0.9)
        vista_detalle_fecha_v4 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/vista_detalles_fecha_V4.png', grayscale = True,confidence=0.9)
        vista_detalle_fecha_v5 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/vista_detalles_fecha_V5.png', grayscale = True,confidence=0.9)
        vista_detalle_fecha_v6 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/vista_detalles_fecha_V6.png', grayscale = True,confidence=0.9)
        if crmAttempts == 6:
            crmAttempts = 0   
            pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
            sleep(0.5) 
            return 10 #returning special number for specific error scenario
        sleep(0.5)
        crmAttempts += 1
        print(crmAttempts)


        sleep(1)            
    if vista_detalle_fecha_v1 != None:
        print('vista_detalle_fecha_v1 detected on screen')        
        pressingKey('tab',2)    
        sleep(0.5)
        selectToEnd()
        pyautogui.write(formatDate(fechaCompromiso))
        sleep(0.5)
        pressingKey('tab',5)
        sleep(0.5)
        selectToEnd()
        sleep(0.5)
        pyautogui.write(formatDate(fechaCompromiso))
        sleep(1)
    elif vista_detalle_fecha_v2 != None:
        print('vista_detalle_fecha_v2 detected on screen')
        pressingKey('tab',4)    
        sleep(0.5)
        selectToEnd()
        pyautogui.write(formatDate(fechaCompromiso))
        sleep(0.5)
        pressingKey('tab',12)
        sleep(0.5)
        selectToEnd()
        sleep(0.5)
        pyautogui.write(formatDate(fechaProgramacion))
        sleep(1)
    elif vista_detalle_fecha_v4 != None or vista_detalle_fecha_v5 != None:

        print('vista_detalle_fecha_v4 or vista_detalle_fecha_v5 or vista_detalle_fecha_v6 detected on screen')
        pressingKey('tab',1)
        sleep(0.5)
        selectToEnd()
        pyautogui.write(formatDate(fechaCompromiso))
        sleep(0.5)
        pressingKey('tab',4)
        sleep(0.5)
        selectToEnd()
        sleep(0.5)
        pyautogui.write(formatDate(fechaProgramacion))
        sleep(1)        
    elif vista_detalle_fecha_v6 != None:
        print('vista_detalle_fecha_v6 detected on screen')
        pressingKey('tab',1)
        sleep(0.5)
        selectToEnd()
        pyautogui.write(formatDate(fechaCompromiso))
        sleep(0.5)
        pressingKey('tab',4)
        sleep(0.5)
        selectToEnd()
        sleep(0.5)
        pyautogui.write(formatDate(fechaProgramacion))
        sleep(1)        
    elif vista_detalle_fecha_v3 != None:
        print('vista_detalle_fecha_v3 detected on screen')
        pressingKey('tab',1)    
        sleep(0.5)
        selectToEnd()
        pyautogui.write(formatDate(fechaCompromiso))
        sleep(0.5)
        pressingKey('tab',3)
        sleep(0.5)
        selectToEnd()
        sleep(0.5)
        pyautogui.write(formatDate(fechaProgramacion))
        sleep(1)
    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)
    crmAttempts = 0
    while crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
                
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario REVISADO")
            print("CRM ASSIGN USER:",crm_assign_user)
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

         # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
        # Close Edit Incident View 
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        return 0
    else:
        sleep(1)
        pressingKey('tab')
        sleep(0.5)
        pressingKey('tab')
        sleep(1)
        pressingKey('enter')
        print("CRM empty_warning_message_crm pop up is present!")
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        
        # You must Validate assign user pop up is visible and on focus
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs         
        print("n key has been pressed!")  
      
    sleep(1)  
    return 1   

def searchAndUpdateBillingItem(incidentId,BillingItem):    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = crm_warning_message = mod_consulta_popup = None
    billingItem_field = aliado_implementacion_field = aliado_implementacion_field2 = None
    crmAttempts = 0
    print("BillingItem: ",BillingItem)
    print("incidentId: ",incidentId)
    
    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
        print(crmAttempts)
        
    print("CRM Dashboard GUI is present!")
    pyautogui.click(pyautogui.center(crm_dashboard))   
    sleep(1)
    pressingKey('f2')
    
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(1)
    pressingKey('enter')
    
    #/////////////////////////////////// EDIT VIEW CRM /////////////////////////////////////////// 
    
    # Validate edit_incident view is visible and on focus    
    while crm_edit_incident is None:
        crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
    print("CRM Edit Incident button is present!")
    crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
    pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)

    # Validate assign user pop up is visible and on focus   
    while crm_assign_user is None:
        crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
    print("CRM Confirm Assign Pop Up is present!")        
    sleep(1)
    pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
    sleep(1)
    
    # Maximize CRM Edit window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    sleep(1)
    pyautogui.click(24,619)
    sleep(1)

    #/////////////////////////////////// UPDATE BILLING ITEM  /////////////////////////////////////  
    # Identify which UI is present on screen. Then focus on BILLING ITEM FIELD and select the incoming billing item 
        
    billingItem_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/items_facturacion_modulo.png', grayscale = True,confidence=0.95)
    billingItem_field2 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/items_facturacion_modulo2.png', grayscale = True,confidence=0.95)
    print(billingItem_field)
    scrollTimes = 0
    pyautogui.moveTo(180, 763)    
    while billingItem_field is None and billingItem_field2 is None and scrollTimes <= 300:
        pyautogui.scroll(-15)
        sleep(0.03)
        billingItem_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/items_facturacion_modulo.png', grayscale = True,confidence=0.95)
        billingItem_field2 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/items_facturacion_modulo2.png', grayscale = True,confidence=0.95)
        scrollTimes += 1
    if(scrollTimes > 299):
        print("billing item field was not found on GUI screen!")    
    else:
        print("billing item field is present and detected on GUI screen!")      
    
    while aliado_implementacion_field is None and aliado_implementacion_field2 is None:
        print("buscando aliado_implementacion_field in screen")
        aliado_implementacion_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/aliado_implementacion.png', grayscale = True,confidence=0.95)        
        aliado_implementacion_field2 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/aliado_implementacion2.png', grayscale = True,confidence=0.95)        
    print("Aliado Implementacion field is present!")            
    if(aliado_implementacion_field2 is None):        
        aliado_implementacion_field_x,aliado_implementacion_field_y = pyautogui.center(aliado_implementacion_field)    
        pyautogui.moveTo(aliado_implementacion_field_x,aliado_implementacion_field_y)
        pyautogui.moveRel(144,0)
    else:    
        aliado_implementacion_field_x,aliado_implementacion_field_y = pyautogui.center(aliado_implementacion_field2)    
        pyautogui.moveTo(aliado_implementacion_field_x,aliado_implementacion_field_y)
        pyautogui.moveRel(144,0)
        
    pyautogui.click()
    pyautogui.write("NAE")
    pressingKey('tab')
    pyautogui.write("NAE")
    pressingKey('tab')
    pyautogui.write("GESTION")    
    pressingKey('tab')
    pyautogui.write(BillingItem)
    pressingKey('tab')
    
    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)
    crmAttempts = 0
    while crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
                
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario REVISADO")
            print("CRM ASSIGN USER:",crm_assign_user)
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

         # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
        # Close Edit Incident View 
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        return 0
    else:
        sleep(1)
        pressingKey('tab')
        sleep(0.5)
        pressingKey('tab')
        sleep(1)
        pressingKey('enter')
        print("CRM empty_warning_message_crm pop up is present!")
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        
        # You must Validate assign user pop up is visible and on focus
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs         
        print("n key has been pressed!")  
      
    sleep(1)  
    return 1  

def searchAndUpdateIncident(incidentId):
    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = crm_warning_message = mod_consulta_popup = None    
    estado_field = usuario_asignado  = cod_resolucion_1_field = None
    crmAttempts = 0    
    print("incidentId: ",incidentId)
    
    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
        print(crmAttempts)
    print("CRM Dashboard GUI is present!")    
    
    pyautogui.click(pyautogui.center(crm_dashboard))   
    sleep(1)
    pressingKey('f2')
    
    while mod_consulta_popup is None:
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(1)
    pressingKey('enter')
    
    #/////////////////////////////////// EDIT VIEW CRM /////////////////////////////////////////// 
    
    # Validate edit_incident view is visible and on focus    
    while crm_edit_incident is None:
        crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
    print("CRM Edit Incident button is present!")
    crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
    pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)

    # Validate assign user pop up is visible and on focus   
    while crm_assign_user is None:
        crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
    print("CRM Confirm Assign Pop Up is present!")        
    sleep(1)
    pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
    sleep(1)
    
    # Maximize CRM Edit window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    sleep(1)
    
    #/////////////////////////////////// UPDATE INCIDENT  /////////////////////////////////////        
    
    # CAMBIO DE ESTADO
    while estado_field is None:
        print("buscando estado_field in screen")
        estado_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/estado_field.png', grayscale = True,confidence=0.95)        
    print("estado_field is present!")            
        
    estado_field_x,estado_field_y = pyautogui.center(estado_field)    
    pyautogui.moveTo(estado_field_x,estado_field_y)
    pyautogui.moveRel(265,0)
    pyautogui.click()
    pyautogui.write("PENDIENTE CLIENTE")
    pressingKey('enter')

    # CAMBIO DE USUARIO
    while usuario_asignado is None:
        print("buscando usuario_asignado in screen")
        usuario_asignado = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/usuario_asignado.png', grayscale = True,confidence=0.95)
    print("usuario_asignado is present!")            
        
    usuario_asignado_x,usuario_asignado_y = pyautogui.center(usuario_asignado)
    pyautogui.moveTo(usuario_asignado_x,usuario_asignado_y)
    pyautogui.moveRel(265,0)     
    pyautogui.click()
    sleep(0.7)    
    pyautogui.write("YULY ANDREA ARDILA SANCHEZ")        
    pressingKey('enter')

    # CAMBIO DE CODIGO DE RESOLUCION 1
    while cod_resolucion_1_field is None:
        print("buscando cod_resolucion_1_field in screen")
        cod_resolucion_1_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/cod_resolucion_1_field.png', grayscale = True,confidence=0.95)        
    print("cod_resolucion_1_field is present!")
    
    cod_resolucion_1_field_x,cod_resolucion_1_field_y = pyautogui.center(cod_resolucion_1_field)    
    pyautogui.moveTo(cod_resolucion_1_field_x,cod_resolucion_1_field_y)
    pyautogui.moveRel(181,0)     
    pyautogui.click()
    pyautogui.write("3-SUSPENDER")
    pressingKey('enter')

    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        print("buscando crm_save_incident in screen")
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)
    crmAttempts = 0
    while crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
                
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario REVISADO")
            print("CRM ASSIGN USER:",crm_assign_user)
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

         # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            print("buscando crm_otp_saved_sucessfully en pantalla")
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
        # Close Edit Incident View 
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        return 0
    else:
        sleep(1)
        pressingKey('tab')
        sleep(0.5)
        pressingKey('tab')
        sleep(1)
        pressingKey('enter')
        print("CRM empty_warning_message_crm pop up is present!")
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        
        # You must Validate assign user pop up is visible and on focus
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs         
        print("n key has been pressed!")  
      
    sleep(1)  
    return 1  
