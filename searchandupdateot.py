import pyautogui
import pyperclip
from time import sleep
from genericfunctions import (pressingKey,selectToEnd,formatDate)

def searchAndUpdateUser(incidentId,ususarioAsignado):    
    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = crm_warning_message = crm_ot_blocked_message = None
    asignado_a = mod_consulta_popup = None 
    crmAttempts = 0
        
    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    pyautogui.click(905,505)
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
                    
    print("CRM Dashboard GUI is present!")   
    pyautogui.click(pyautogui.center(crm_dashboard))
    

    sleep(0.5)
    pressingKey('f2')
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    #/////////////////////////////////// FASE DE ADVERTENCIAS ///////////////////////////////////////////
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")
        crmAttempts += 1
        print(crmAttempts)            
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        
        print("No se encuentrarn ventanas de advertencia!")                                   
        # Validate edit_incident view is visible and on focus
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("Vista de Detalles OT is present!")    
        print("CRM Edit Incident button is present!")
        crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
        pyautogui.click(crm_edit_incident_x,crm_edit_incident_y)
    
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
        pyautogui.click(pyautogui.center(crm_ot_blocked_message))
        pressingKey('enter')
        sleep(2)
        pyautogui.hotkey('alt','f4')
        return 9
    
    #/////////////////////////////////// MAXIMIZE WINDOW /////////////////////////////////////////// 
    
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(24,617)
    crmAttempts = 0
    
    #/////////////////////////////////// UPDATE USER ASSIGNED /////////////////////////////////////  
    # Identify which UI is present on screen. Then focus on USER ASSIGNATION FIELD and Paste the name 
    # of the user passed as parameters
            
    while asignado_a is None:        
        asignado_a = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignado_a.png', grayscale = True,confidence=0.85)
    print("asignado_a field is present and detected on GUI screen!")
    crm_asignado_a_x,crm_asignado_a_y = pyautogui.center(asignado_a)
    pyautogui.moveTo(crm_asignado_a_x,crm_asignado_a_y)
    pyautogui.moveRel(250,0)
    pyautogui.click()
    pyautogui.write(ususarioAsignado)
    pressingKey('tab')
    
    #/////////////////////////////////// CLOSING CRM EDIT/DETAILS VIEW ///////////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)
    sleep(1)
                
    crm_assign_user = None
    while crm_warning_message is None and crm_assign_user is None:  
        print("buscando o crm_warning_message o crm_assign_user")      
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
        sleep(0.5)    
    
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
        print("CRM empty_warning_message_crm pop up is present!")
        pyautogui.click(pyautogui.center(crm_warning_message))
        sleep(1)
        pressingKey('enter')
        sleep(1)
        pressingKey('enter')
        
        # Validate assign user pop up is visible and on focus 
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 10:            
            crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            crmAttempts += 1
        crmAttempts = 0
        
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)        
        
        # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None and crmAttempts < 10:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            crmAttempts += 1
        if(crm_otp_saved_sucessfully is not None):
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

def searchAndUpdateDates(incidentId,fechaProgramacion,fechaCompromiso,anotaciones):
    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = anotaciones_ot_field = None     
    fecha_programacion_field = fecha_programacion_field_2 = fecha_compromiso_field = mod_consulta_popup = None
    crm_warning_message = crm_ot_blocked_message = None
    crmAttempts = 0
    
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
                
    print("CRM Dashboard GUI is present!")       
    pyautogui.click(pyautogui.center(crm_dashboard))    
    
    #/////////////////////////////////// VENTANA DE CONSUlTA ///////////////////////////////////////////
    
    sleep(0.5)
    pressingKey('f2')
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    #/////////////////////////////////// FASE DE ADVERTENCIAS ///////////////////////////////////////////
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        print("No se encuentran ventanas de advertencia!")                           
        
        # Validate edit_incident view is visible and on focus            
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("Vista de Detalles is present!")    
        print("CRM Edit Incident button is present!")
        crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
        pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
    
        # Validate assign user pop up is visible and on focus
        # crm_assign_user = None  # reset variable   
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 10:
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
        pressingKey('enter')
        sleep(3)
        pressingKey('enter')
        sleep(1)
        pyautogui.hotkey('alt','f4')
        return 9
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")
    
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(24,617)
    crmAttempts = 0
    
    #/////////////////////////////////// ACTUALIZACION DE FECHAS /////////////////////////////////////  
    # move to specific area
    
    pyautogui.moveRel(250,0)
    # CAMBIO DE FECHA DE PROGRAMACION                
    while fecha_programacion_field is None and fecha_programacion_field_2 is None and crmAttempts < 250:
        print("buscando fecha_programacion_field in screen")
        print("crmAttempts: ",crmAttempts)
        fecha_programacion_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/fecha_programacion_field.png', grayscale = True,confidence=0.95)        
        fecha_programacion_field_2 = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/fecha_programacion_field_2.png', grayscale = True,confidence=0.95)                    
        if crmAttempts % 5  == 0 and crmAttempts > 20:
            pyautogui.scroll(-35)        
        crmAttempts = crmAttempts + 1 
    
    if crmAttempts > 249:        
        return 9
            
    if fecha_programacion_field_2 is None:        
        pyautogui.moveTo(pyautogui.center(fecha_programacion_field))                    
    else:
        pyautogui.moveTo(pyautogui.center(fecha_programacion_field_2))          
        
    pyautogui.click()    
    pyautogui.moveRel(250,0)    
    pyautogui.click()
    pyautogui.hotkey('fn','end')  
    pyautogui.hotkey('fn','shift','home')      
    pyautogui.write(formatDate(fechaProgramacion))
    pressingKey('enter')
    
    # ----- CAMBIO DE FECHA DE COMPROMISO ------
    crmAttempts = 1
    while fecha_compromiso_field is None and  crmAttempts < 300:
        print("buscando fecha_compromiso_field in screen")
        fecha_compromiso_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/fecha_compromiso_field.png', grayscale = True,confidence=0.95)                
        
        if crmAttempts % 5 == 0:
            pyautogui.scroll(-35)            
        crmAttempts = crmAttempts + 1                 
    
    print("fecha_compromiso_field is present!")
    pyautogui.moveTo(pyautogui.center(fecha_compromiso_field))            
    pyautogui.click()
    pyautogui.moveRel(250,0)
    pyautogui.click()        
    pyautogui.hotkey('fn','end')  
    pyautogui.hotkey('fn','shift','home')          
    pyautogui.write(formatDate(fechaCompromiso))    
    pressingKey('enter')
    
    #/////////////////////////////////// TRASLADO DE COMENTARIOS /////////////////////////////////////
    if(anotaciones != "None"):    
    
        while anotaciones_ot_field is None:
            print("buscando anotaciones_ot_field in screen")
            anotaciones_ot_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/anotaciones_ot_field.png', grayscale = True,confidence=0.95)                                
        
        print("anotaciones_ot_field is present!")
        pyautogui.moveTo(pyautogui.center(anotaciones_ot_field))            
        pyautogui.click()
        pressingKey('tab')
        sleep(0.5)
        pressingKey('tab')
        pyperclip.copy(anotaciones)
        sleep(0.5)
        pyautogui.hotkey('ctrl','v')    
        sleep(0.5)
    
    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)    
        
    crm_assign_user = None 
    crmAttempts = 0
    while crm_warning_message is None and crmAttempts < 5:  
        print("buscando crm_warning_message")      
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)              
        sleep(0.5)
        crmAttempts +=1
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
                
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 10:
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
            pass

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
        print("CRM empty_warning_message_crm pop up is present!")
        #pyautogui.click(pyautogui.center(crm_warning_message))
        sleep(0.5)
        pressingKey('enter')
        sleep(1)
        pressingKey('enter')
        sleep(1)
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1)
        pressingKey('n')        
            
        return 9
        
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
                
    print("CRM Dashboard GUI is present!")   
    pyautogui.click(pyautogui.center(crm_dashboard))    
    sleep(0.5)
    pressingKey('f2')

    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    # ---- FASE DE ADVERTENCIAS ----
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        print("No se encuentrarn ventanas de advertencia!")                           
        
        # Validate edit_incident view is visible and on focus            
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("Vista de Detalles is present!")    
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
    
    # ------- ACCIONES FASE DE ADVERTENCIA --------
    else:
        pressingKey('enter')
        sleep(3)
        pressingKey('enter')
        sleep(1)
        pyautogui.hotkey('alt','f4')
        return 9
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")    
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(24,617)
    crmAttempts = 0

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

def searchAndUpdateState(incidentId,estado,usuario,cod_resolucion_1):
    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = crm_warning_message = mod_consulta_popup = crm_ot_blocked_message = None
    estado_field = usuario_asignado  = cod_resolucion_1_field = crm_correo_enviado = None
    crmAttempts = 0
    estados = estado.split(",", 2) # se divide la cadena de texto con los estados a gestionar
    usuarios = usuario.split(",", 2) # se divide la cadena de texto con los usuarios que deben ser asignados
    print(estados,usuarios)
    
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

    print("CRM Dashboard GUI is present!")
    pyautogui.click(pyautogui.center(crm_dashboard))    
    sleep(0.5)
    pressingKey('f2')

    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    # ---- FASE DE ADVERTENCIAS ----
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)  
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        print("No se encuentrarn ventanas de advertencia!")                           
        
        # Validate edit_incident view is visible and on focus            
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("Vista de Detalles is present!")    
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
    
    # ------- ACCIONES FASE DE ADVERTENCIA --------
    else:
        pressingKey('enter')
        sleep(3)
        pressingKey('enter')
        sleep(1)
        pyautogui.hotkey('alt','f4')
        return 9
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(24,617)
    crmAttempts = 0
    
    #/////////////////////////////////// CAMBIO DE ESTADO FLUJO #1 UPDATE STATE  /////////////////////////////////////
    
    # ------------ 1ER CICLO ------------#
    # CAMBIO DE ESTADO
    while estado_field is None:
        print("buscando estado_field in screen")
        estado_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/estado_field.png', grayscale = True,confidence=0.95)        
    print("estado_field is present!")
    
    estado_field_x,estado_field_y = pyautogui.center(estado_field)    
    pyautogui.moveTo(estado_field_x,estado_field_y)
    pyautogui.moveRel(265,0)
    pyautogui.click()
    pyautogui.write(estados[0])
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
    pyautogui.write(usuarios[0])        
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
    pyautogui.hotkey('fn', 'home')
    pyautogui.keyDown('enter')
    sleep(1)    
    pyautogui.write(cod_resolucion_1)
    pressingKey('enter')
  
    while crm_save_incident is None:
        print("buscando crm_save_incident in screen")
        crm_save_incident = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    sleep(0.5)
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
                
        # ------------ 2ND CYCLE -------------------#
        # MAKE IT DINAMYC VALIDATION IF NECESSARY
        # ---- CAMBIO DE ESTADO --------        
        if len(estados) > 1:
            while estado_field is None:
                print("buscando estado_field 2do intento in screen")
                estado_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")            
            
            estado_field_x,estado_field_y = pyautogui.center(estado_field)    
            pyautogui.moveTo(estado_field_x,estado_field_y)
            pyautogui.moveRel(265,0)
            pyautogui.click()
            pyautogui.write(estados[1])
            pressingKey('enter')

            # CAMBIO DE USUARIO
            while usuario_asignado is None:
                print("buscando usuario_asignado in screen")
                usuario_asignado = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/usuario_asignado.png', grayscale = True,confidence=0.95)
            print("usuario_asignado 2do intento is present!")            
                
            usuario_asignado_x,usuario_asignado_y = pyautogui.center(usuario_asignado)
            pyautogui.moveTo(usuario_asignado_x,usuario_asignado_y)
            pyautogui.moveRel(265,0)     
            pyautogui.click()
            sleep(0.7)
            pyautogui.write(usuarios[1])        
            pressingKey('enter')

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
                
                if len(estados) == 2:
                    crmAttempts = 0                    
                    while crm_correo_enviado is None  and crmAttempts < 90:
                        print("buscando crm_correo_enviado en pantalla. intento # ",crmAttempts)
                        sleep(1)
                        crm_correo_enviado = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/correo_enviado_cambio_estado.png', grayscale = True,confidence=0.9)
                        crmAttempts += 1
                    sleep(0.15)
                    if crmAttempts > 88:
                        return 9
                    pyautogui.click(pyautogui.center(crm_correo_enviado))
                    pressingKey('tab')
                    pressingKey('enter')
                
                # Validate assign user pop up is visible and on focus   
                crm_assign_user = None  # reset variable            
                while crm_assign_user is None:
                    crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
                    sleep(0.5)
                    print("buscando ventana de asignación de usuario REVISADO")                                
                print("CRM Confirm Assign Pop Up is present!")
                sleep(0.5)
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

        # ------------ 3ER CICLO -------------------#
        # MAKE IT DINAMYC VALIDATION IF NECESSARY
        # ---- CAMBIO DE ESTADO --------      
        if len(estados) > 2:
            while estado_field is None:
                print("buscando estado_field 2do intento in screen")
                estado_field = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")            
            
            estado_field_x,estado_field_y = pyautogui.center(estado_field)    
            pyautogui.moveTo(estado_field_x,estado_field_y)
            pyautogui.moveRel(265,0)
            pyautogui.click()
            pyautogui.write(estados[2])
            pressingKey('enter')

            # CAMBIO DE USUARIO
            while usuario_asignado is None:
                print("buscando usuario_asignado in screen")
                usuario_asignado = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/usuario_asignado.png', grayscale = True,confidence=0.95)
            print("usuario_asignado 2do intento is present!")            
                
            usuario_asignado_x,usuario_asignado_y = pyautogui.center(usuario_asignado)
            pyautogui.moveTo(usuario_asignado_x,usuario_asignado_y)
            pyautogui.moveRel(265,0)     
            pyautogui.click()
            sleep(0.7)
            pyautogui.write(usuarios[1])        
            pressingKey('enter')
            
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
                crmAttempts = 0                    
                while crm_correo_enviado is None  and crmAttempts < 90:
                    print("buscando crm_correo_enviado en pantalla. intento # ",crmAttempts)
                    sleep(1)
                    crm_correo_enviado = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/correo_enviado_cambio_estado.png', grayscale = True,confidence=0.9)
                    crmAttempts += 1
                sleep(0.15)
                if crmAttempts > 88:
                    return 9
                pyautogui.click(pyautogui.center(crm_correo_enviado))
                pressingKey('tab')
                pressingKey('enter')
                
                # Validate assign user pop up is visible and on focus   
                crm_assign_user = None  # reset variable            
                while crm_assign_user is None:
                    crm_assign_user = pyautogui.locateOnScreen('C:/BOT BPO Automation/Version 1.0/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
                    sleep(0.5)
                    print("buscando ventana de asignación de usuario REVISADO")                                
                print("CRM Confirm Assign Pop Up is present!")
                sleep(0.5)
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
