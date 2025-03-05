import os
import time
import win32com.client
import pandas as pd

# =========================
# Configuración de Rutas
# =========================
FOLDER_PATH = r"C:\TEMP\Backup Existencias"
MB52_FILE = os.path.join(FOLDER_PATH, "MB52.xlsx")

# =========================
# Funciones de SAP
# =========================
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



def ejecutar_MB52(session):
    """
    Ejecuta la transacción MB52 en SAP utilizando los IDs proporcionados y exporta el reporte a MB52.XLSX.
    """
    try:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "MB52"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1000"
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").caretPosition = 4
        session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "PB*"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").caretPosition = 3
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]/usr/ctxtMATKLA-LOW").text = "C328"
        session.findById("wnd[0]/usr/ctxtP_VARI").text = "/mb52 avon"
        session.findById("wnd[0]/usr/ctxtP_VARI").setFocus()
        session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\TEMP\\Backup Existencias"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB52.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)

        print("Transacción MB52 ejecutada y exportada a:", MB52_FILE)
    except Exception as e:
        print("Error durante la ejecución de MB52:", e)
        
        time.sleep(5)
        

def cerrar_excel(tiempo_espera=2):
    """
    Espera un tiempo determinado y cierra la instancia de Excel abierta.
    Esto es útil para evitar que queden procesos abiertos que bloqueen la edición de los archivos.
    """
    try:
        print(f"Esperando {tiempo_espera} segundos antes de cerrar Excel...")
        time.sleep(tiempo_espera)
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            workbook = excel.Workbooks(os.path.basename(MB52_FILE))
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
        count = excel.Workbooks.Count
        if count > 0:
            for i in range(count, 0, -1):
                excel.Workbooks[i].Close(SaveChanges=False)
            print("Se cerraron todos los libros abiertos en Excel.")
        excel.Quit()
        print("Excel se cerró automáticamente.")
    except Exception as e:
        print("No se pudo cerrar Excel, probablemente no estaba abierto.", e)

        

# =========================
# Función Principal
# =========================
def main():
    session = conectar_sap()
    if session is None:
        return
    
    ejecutar_MB52(session)
    
    cerrar_excel(tiempo_espera=2)


if __name__ == "__main__":
    main()
