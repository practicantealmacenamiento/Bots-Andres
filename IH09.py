import os
import time
import win32com.client
import pandas as pd

# =========================
# Configuración de Rutas
# =========================
FOLDER_PATH = r"C:\TEMP\Backup Existencias"
IH09_FILE = os.path.join(FOLDER_PATH, "IH09.xlsx")

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

def ejecutar_IH09(session):
    """
    Ejecuta la transacción IH09 en SAP utilizando los IDs proporcionados y exporta el reporte a IH09.XLSX.
    """
    try:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "ih09"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").text = "mp*"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(12, "MAKTG")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/chkCB_ALWAYS").setFocus()
        session.findById("wnd[1]/usr/chkCB_ALWAYS").selected = True
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\TEMP\\Backup Existencias"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "IH09.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)


        print("Transacción IH09 ejecutada exitosamente.")
    except Exception as e:
        print(f"Error al ejecutar IH09: {e}")
        
def close_excel(tiempo_espera=1):
    """
    Espera un tiempo determinado y cierra la instancia de Excel abierta.
    Esto es útil para evitar que queden procesos abiertos que bloqueen la edición de los archivos.
    """
    try:
        print(f"Esperando {tiempo_espera} segundos antes de cerrar Excel...")
        time.sleep(tiempo_espera)
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            workbook = excel.Workbooks(os.path.basename(IH09_FILE))
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
    
    ejecutar_IH09(session)
    
    time.sleep(2)
    close_excel(tiempo_espera=1)


if __name__ == "__main__":
    main()
