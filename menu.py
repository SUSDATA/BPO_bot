def show():

    print("----------------- BOT BPO v1.0 --------------------")                        
    print("Introduzca el numero de la actividad que desea ejecutar")
    actividad_bot = input("1. Cambiar Fechas\n2. Cambiar Usuario\n3. Items de Facturación\n4. Cambio de Estado (Flujo Instalación - pase a entrega final)\n5. Cambio de Estado (Flujo Novedades - pase a entrega final)\n")
    act_bot = int(actividad_bot)
    
    if act_bot > 5:
        return -1
    else:
        return act_bot        
    