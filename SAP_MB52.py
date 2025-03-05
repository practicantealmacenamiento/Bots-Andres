import os
import time
import win32com.client

# Configuración de destino
FOLDER_PATH = r"C:\TEMP\Informe seguros"
EXCEL_FILE = os.path.join(FOLDER_PATH, "Informe_MB52.xlsx")

def conectar_sap():
    """Conecta a SAP y obtiene la sesión activa."""
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        if application.Children.Count == 0:
            raise Exception("No hay conexiones activas. Conéctate manualmente en SAP Logon.")
        
        connection = application.Children(0)
        session = connection.Children(0)
        
        print("Conectado a SAP exitosamente.")
        return session
    except Exception as e:
        print(f"Error al conectar a SAP: {e}")
        return None
    

def ejecutar_exp(session):
    """Ejecuta la transacción MB52 y exporta el archivo sin ejecutarlo."""
    session.findById("wnd[0]").maximize()  # Maximiza la ventana principal de SAP
    session.findById("wnd[0]/tbar[0]/okcd").text = "mb52"  # Ingresa la transacción MB52 en la barra de comandos
    session.findById("wnd[0]").sendVKey(0)  # Presiona Enter para ejecutar la transacción
    
    # Limpia el campo de ubicación de almacenamiento
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").setFocus()  # Enfoca el campo de ubicación
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 0  # Posiciona el cursor al inicio
    session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press()  # Presiona el botón de selección múltiple
    
    # Selecciona varias ubicaciones de almacenamiento
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1002"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1003"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1008"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "1011"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "1014"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "1015"
    
    # Posiciona el cursor en la última celda seleccionada
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 4
    
    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Presiona el botón "Continuar"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Ejecuta la consulta en MB52
    
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()  # Accede a la opción de exportar
    
    # Configura la ruta y el nombre del archivo de exportación
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\TEMP\\Informe seguros"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Informe_MB52.xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17  # Posiciona el cursor al final del nombre de archivo
    
    session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guarda el archivo
    
    # Cierra las ventanas y regresa a la pantalla inicial
    session.findById("wnd[0]").sendVKey(3)  # Presiona "Atrás"
    session.findById("wnd[0]").sendVKey(3)  # Presiona "Atrás" nuevamente

    


def cerrar_excel(tiempo_espera=2):
    """
    Espera un tiempo determinado (en segundos) antes de cerrar Excel.
    Esto permite que Excel complete cualquier operación pendiente o que se abra el libro,
    para luego forzar su cierre.
    """
    try:
        print(f"Esperando {tiempo_espera} segundos antes de cerrar Excel...")
        time.sleep(tiempo_espera)
        # Conectar con la instancia de Excel
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks("Informe_MB52.xlsx")
        workbook.Close(SaveChanges=False)
        count = excel.Workbooks.Count
        if count > 0:
            for i in range(count, 0, -1):
                excel.Workbooks[i].Close(SaveChanges=False)
            print("Se cerraron todos los libros abiertos en Excel.")
        excel.Quit()
        print("Excel se cerró automáticamente.")
    except Exception as e:
        print("No se pudo cerrar Excel, probablemente no estaba abierto.", e)



            
    
# Ejecución del script
session = conectar_sap()

if session:
    ejecutar_exp(session)  
    time.sleep(5)  # Esperar unos segundos para asegurar que Excel tenga tiempo de abrir el archivo  
    cerrar_excel() # cerramos Excel después de generar el informe  

