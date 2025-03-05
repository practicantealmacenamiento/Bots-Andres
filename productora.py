import os
import time
import win32com.client

# Configuración de destino
FOLDER_PATH = r"C:\TEMP\Backup Existencias"
EXCEL_FILE = os.path.join(FOLDER_PATH, "Backup_Medellin_Rionegro.xlsx")

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
    
def cerrar_excel(tiempo_espera=1):
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
        workbook = excel.Workbooks("Backup_Medellin_Rionegro.xlsx")
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
        

def ejecutar_exp(session):
    """Ejecuta la transacción CED y exporta el archivo sin ejecutarlo."""
    try:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "LX02"
        session.findById("wnd[0]").sendVKey(0)

        # Ingresar parámetros
        session.findById("wnd[0]/usr/ctxtS1_LGNUM").text = "PRO"
        session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtP_VARI").text = "/WM_ALMACEN"

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(5)  # Esperar que cargue la transacción

        print("LX02 ejecutada correctamente.")
        
        # Exportar archivo: abrir menú de exportación
        session.findById("wnd[0]").sendVKey(16)  # Llamar menú de exportación
        time.sleep(1)

        # Establecer ruta y nombre de archivo
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FOLDER_PATH
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Backup_Medellin_Rionegro.xlsx"

        # Verificar si el archivo ya existe y eliminarlo
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
            print("Archivo existente eliminado.")

        # Confirmar guardado (sin ejecutar el archivo)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(5)  # Esperar a que se complete la exportación


       # Cerrar la transacción en SAP enviando F3 (equivalente a sendVKey(3))
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)

        
        time.sleep(1)
        
        # Cerrar Excel si se abrió automáticamente
        cerrar_excel()

        print(f"Archivo exportado correctamente a: {EXCEL_FILE}")

    except Exception as e:
        print(f"Error al ejecutar PRO o exportar: {e}")

# Ejecución del script
session = conectar_sap()

if session:
    ejecutar_exp(session)
