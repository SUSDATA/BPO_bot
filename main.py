
# VALIDATE CRM DASHBOARD IS IN FOCUS AND SEARCH FOR A NON IMAGE DETECTION SOLUTION
# while crm_dashboard is None:
#         crm_dashboard = pyautogui.locateOnScreen('crm_dashboard.png', grayscale = True,confidence=0.85)    
#     print("CRM Dashboard GUI is present!")

import pyautogui
import asyncio
import sys
import os
import openpyxl # to handle input data through Excel
from openapp import (openCRM,connectToCRM,openAndConnectVPNForticlient)
from searchandupdateot import (formatDate,searchAndUpdateDates,searchAndUpdateUser,searchAndUpdateBillingItem)
from closeapps import (closeCRM,disconnectVPNForticlient,closeFortiClient)
from genericfunctions import (showDesktop,make_noise,get_username_os,sendEmail,getCurrentDateAndTime)
from telegramfunctions import (sendTelegramMsg,sendTelegramMsgWithDocuments)
import logging
#from logsfunctions import (setupLogger,setupLogger_2)
from time import sleep

async def main():
            
    # CONSTANTS
    TELEGRAM_CHAT_ID = 5970685607 # PERSONAL CHAT WITH BOT 5970685607 # TELEGRAM CHAT GROUP ID #-1002019721248
    INPUT_DIRECTORY = 'C:/BOT INSUMO BASE'
    INPUT_FILENAME = 'Input BOT 001 V1.xlsx'
    SUPER_LOG_FILENAME = 'C:/BOT BPO Automation/Version 1.0/logs/super_log.txt'
    #INIT_LOG_FILENAME = 'logs/init_log.txt'
    CORREOS_DESTINO = 'daniel.garcia@nae.global;danielgarcc@gmail.com;catherine.vargas@nae.global;hladb@nae.global;dbenv@nae.global'
    CORREOS_CC = 'drobinic.daniel@gmail.com'

    # VARIABLES        
    total_rows = 0
    ot_completadas = []   
    resp_funcion = None
        
    try:
        #***************************************#
        #******* LOGGING CONFIGURATION *********#
        #***************************************#
        
        # SUPER LOGGER
        FORMAT ='%(asctime)s - %(name)s - %(user)s - %(levelname)s - %(message)s'
        logging.basicConfig(filename=SUPER_LOG_FILENAME, filemode='w',format=FORMAT, datefmt='%m/%d/%Y %I:%M:%S %p',level=logging.WARNING)
        attrs = {'user': get_username_os()}

        """To setup as many loggers as you want"""
        # formatter = logging.Formatter('%(user)s - %(name)s - %(levelname)s - %(message)s')
                
        # handler = logging.FileHandler(INIT_LOG_FILENAME)        
        # handler.setFormatter(formatter)

        # init_logger = logging.getLogger('init log')
        # init_logger.setLevel(logging.INFO)
        # init_logger.addHandler(handler)
        
        # sleep(1)
        # logging.warning('LOGGERS have been properly setup',extra=attrs) 
        # logging.warning('Main Function started',extra=attrs) 
        
        #*******************************************#
        #*** SHOW MENU TO SELECT RPA ACTIVITY ******#
        #*******************************************#
        
        print("----------------- RPA BPO v1.0 --------------------")                        
        print("Introduzca el numero de la actividad que desea ejecutar")
        actividad_rpa = input("1. Actualizar Fechas\n2. Asignar Usuario\n3. Items de Facturación\n")
            
        if actividad_rpa == "1":
            actividad_rpa_selected = 'ACTUALIZACIÓN DE FECHAS'
            
        elif  actividad_rpa == "2":
            actividad_rpa_selected = 'ASIGNACIÓN DE USUARIO'    
            
        else:
            print("Actividad no permitida!")            
            print("Exit")
            exit_program()
        logging.warning('Inicio de ejecucion del BOT para %s ',actividad_rpa_selected,extra=attrs)
        
        #*******************************************#
        #************* DATA EXTRACTION *************#
        #*******************************************#
        # Extract Data from Excel File and Open an 
        # existing Excel wb file 
        
        file = os.path.join(INPUT_DIRECTORY,INPUT_FILENAME)
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        if 'Input Backup' in wb.sheetnames:
            del wb['Input Backup']
        
        # Notify the start of RPA execution via Telegram        
        await sendTelegramMsg('START - RPA ' + actividad_rpa_selected + ' ha iniciado.\nUsuario ['+ get_username_os() +']\nTiempo de Inicio: ' + getCurrentDateAndTime(),TELEGRAM_CHAT_ID)
        await sendTelegramMsgWithDocuments(TELEGRAM_CHAT_ID)
        
        # # Notify the start of RPA execution via Email         
        # sendEmail(
        #     'START - RPA ' + actividad_rpa_selected + ' ha iniciado',
        #     CORREOS_DESTINO,
        #     CORREOS_CC,
        #     'START - RPA ' + actividad_rpa_selected + ' ha sido iniciado por el usuario ['+ get_username_os() +'] a las: ' + getCurrentDateAndTime(),
        #     'C:\\BOT INSUMO BASE\\Input BOT 001 V1.xlsx',
        #     'null'            
        # )             
        
        # Excel Input Base Manipulation         
        target_ws = wb.copy_worksheet(ws)
        target_ws.title = "Input Backup"
        total_rows = len(ws['A'])
        print("Total de registros en base fuente: ",total_rows-1)        
        logging.warning('%s registros en base fuente',total_rows-1,extra=attrs)
        logging.warning('Extract Data Process finished',extra=attrs)                     
        
        # Validate if all elements are present 
        # in the GUI and all apps are initializated 
        
        #*************************************#
        #************* OPEN APPS *************#
        #*************************************#
        
        #FortiClient
        # openAndConnectVPNForticlient()
        # # #CRM
        sleep(1)
        openCRM()
        sleep(1)
        connectToCRM()
        logging.warning('Opening APPs Process finished has finished',extra=attrs)        
        
        #*******************************************#
        #************* DATA PROCESSING *************#
        #*******************************************#
        #pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.2]")[0].maximize()    
        #Looping with iter_rows method through all the OTs (registers in the Excel)
        
        rows = ws.iter_rows(min_row=2, max_row=total_rows, min_col=1, max_col=7)
        # #for a,b,c,d,e,f,g in rows:
        logging.warning('-------------------------------',extra=attrs)
        for row in rows:
            # se omiten aquellos registros que en la col STATUS tenga COMPLETADO o contenga la palabra ERROR
            if row[1].value == 'COMPLETADO' or str(row[1].value).find("ERROR") != -1:
                continue                

            #Interacting with the CRM Depending on the user input call the corresponding method            
            if actividad_rpa.lower() == "1":
                print('Validando registro: ',str(row[0].value),row[3].value,row[4].value)
                resp_funcion = searchAndUpdateDates(str(row[0].value),row[3].value,row[4].value)
                print("Respuesta de la función:",resp_funcion)
                
            if actividad_rpa.lower() == "2":
                resp_funcion = searchAndUpdateUser(str(row[0].value),str(row[6].value))  
                print("Respuesta de la función:",resp_funcion)

            #searchAndUpdateBillingItem()       
            
            # Manejo de las posibles respuestas identificadas en los metodos de control
            if resp_funcion == 0:
                ot_completadas = [str(row[0].value)]                
                row[1].value = 'COMPLETADO'
                logging.warning('Se ha procesado el incidente: %s',row[0].value,extra=attrs)
            elif resp_funcion == 10:
                logging.warning('No se identifica vista de detalles de OT reconocida en el CRM',extra=attrs)
                row[1].value = 'ERROR - No se identifica vista de detalles de OT reconocida en el CRM'
            else:                                
                row[1].value = 'ERROR - No Se ha procesado el incidente - Para ver mas detalles ver el archivo de logs'
                logging.warning('No Se ha procesado el incidente: %s',row[0].value,extra=attrs)
            wb.save(file) 
                       
        print("Fin del Proceso Macro")    
        logging.warning('-------------------------------',extra=attrs)     

    except FileNotFoundError as e:
        
        logging.warning('FileNotFoundError Occurred: %s',e,extra=attrs)
        await sendTelegramMsg('END - Error de Ejecución: Archivo Fuente No Encontrado - RPA ' + actividad_rpa_selected + '.',TELEGRAM_CHAT_ID)
        # sendEmail(
        #     'END - Error de Ejecución: FileNotFoundError - RPA ' + actividad_rpa_selected,
        #     CORREOS_DESTINO,
        #     CORREOS_CC,        
        #     'END - Ha ocurrido un error en la ejecución del ' + actividad_rpa_selected + '.\nNo se encuentra el archivo base fuente\n' + 'Usuario ['+ get_username_os() +']\nTiempo de finalización:'+ getCurrentDateAndTime(),
        #     'null',
        #     'null'
        # )
        exit_program()
        
    except Exception as e:
        
        print(f"An error occurred: {e}")
        logging.warning('Error Occurred: %s',e,extra=attrs)
        logging.warning('Error Occurred v2: %s',sys.exc_info()[0],extra=attrs)        
        await sendTelegramMsg('END - Error de Ejecución: ' + str(e) + ' - RPA ' + actividad_rpa_selected + '.\n'+'Tipo de Error: '+str(sys.exc_info()[0]),TELEGRAM_CHAT_ID)
        # sendEmail(
        #     'END - Error de Ejecución: ' + e + ' - RPA ' + actividad_rpa_selected,
        #     CORREOS_DESTINO,
        #     CORREOS_CC,           
        #     'END - Ha ocurrido un error en la ejecución del RPA actividad_rpa_selected.\nError de Ejecución: ' + e + '\nUsuario ['+ get_username_os() +']\nTiempo de finalización:'+ getCurrentDateAndTime(),
        #     'null',
        #     'null'
        # )
        exit_program()
    else:
        #*******************************************#
        #******** IF NO ERRORS OCCURRED ************#
        #*******************************************#
        #***************************************#
        #************* CLOSE APPS  *************#
        #***************************************#
        
        # closeCRM()
        # sleep(1)
        # # disconnectVPNForticlient() 
        # #closeFortiClient()
        # showDesktop()                            

        #***************************************#
        #*********** SEND NOTIFICATIONS  *******#
        #***************************************#

        # Esto se ejecutara si el bloque try se ejecuta sin errores
        print("Try block succesfully executed")
        logging.warning('Main Function finished naturally',extra=attrs)
        await sendTelegramMsg('END - RPA '+actividad_rpa_selected+' ha finalizado.\n' + 'Usuario ['+ get_username_os() +']\nTiempo de finalización:'+ getCurrentDateAndTime(),TELEGRAM_CHAT_ID)
        await sendTelegramMsgWithDocuments(TELEGRAM_CHAT_ID)        
        # sendEmail(
        #     'END - RPA '+ actividad_rpa_selected +' ha finalizado',
        #     CORREOS_DESTINO,
        #     CORREOS_CC,        
        #     'END - RPA ' + actividad_rpa_selected + ' ha finalizado. \n' + 'Usuario ['+ get_username_os() +']\n Tiempo de finalización:'+ getCurrentDateAndTime(),
        #     'C:\\BOT INSUMO BASE\\Input BOT 001 V1.xlsx',
        #     'null'
        # )       
        exit_program()        
    finally:        
        
        sleep(1)
        make_noise() 

        
def exit_program():
    sys.exit(0)    
    
if __name__ == "__main__":
   asyncio.run(main())