
import os
import time
import win32com.client
import datetime

# ────────────────────────────────────
# CONFIGURACIÓN
# ────────────────────────────────────

FOLDER_PATH = r"C:\TEMP\Backup Existencias"
EXCEL_FILE = os.path.join(FOLDER_PATH, "Backup_COOISPI.xlsx")
BASE_DATE = "20250207"  # Este parámetro se usa en get_date_range() si lo requieres; en este ejemplo se ignora al usar la fecha actual

# Asegurarse de que la carpeta de destino existe
if not os.path.exists(FOLDER_PATH):
    os.makedirs(FOLDER_PATH)

# ────────────────────────────────────
# FUNCIONES                          
# ────────────────────────────────────

def get_date_range(formato="%Y%m%d"):
    """
    Calcula el rango de fechas basado en la fecha actual:
      - Fecha inicial: 1 día antes de hoy.
      - Fecha final: 8 días después de la fecha inicial.
    Ejemplo:
      Si hoy es 07.02.2025, se obtendrá:
         Fecha inicial = 06.02.2025
         Fecha final   = 14.02.2025
    """
    today = datetime.date.today()
    fecha_inicial = today - datetime.timedelta(days=1)
    fecha_final = fecha_inicial + datetime.timedelta(days=8)
    return fecha_inicial.strftime(formato), fecha_final.strftime(formato)

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

def cerrar_excel(tiempo_espera=2):
    """
    Espera un tiempo determinado y cierra la instancia de Excel que tenga abierto el archivo exportado.
    Se intenta obtener la instancia activa de Excel; si no se encuentra, se crea una.
    Luego se cierra el libro 'Backup_COOISPI.xlsx' (si está abierto) y se finaliza Excel.
    """
    try:
        print(f"Esperando {tiempo_espera} segundos antes de cerrar Excel...")
        time.sleep(tiempo_espera)
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")
        try:
            workbook = excel.Workbooks("Backup_COOISPI.xlsx")
            workbook.Close(SaveChanges=False)
            print("Se cerró el libro 'Backup_COOISPI.xlsx'.")
        except Exception:
            print("El libro 'Backup_COOISPI.xlsx' no se encontró abierto.")
        # Cerrar cualquier otro libro abierto
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
    """
    Ejecuta la transacción COOISPI, exporta el archivo y luego navega a LX02.
    Utiliza la función get_date_range() para calcular el rango de fechas de forma automática.
    """
    try:
        # Calcular el rango de fechas
        fecha_inicial, fecha_final = get_date_range()
        print(f"Fechas calculadas: Fecha inicial = {fecha_inicial}, Fecha final = {fecha_final}")
        
        # --- Configuración y ejecución de la transacción COOISPI ---
        session.findById("wnd[0]").maximize()
        # Seleccionar el nodo adecuado (ajusta el ID según tu Scripting Tracker)
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00014")
        
        # Configurar parámetros de la transacción
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/DESRIO"
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "LIB."
        
        # Ingresar la fecha inicial
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").setFocus()
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").caretPosition = 0
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = fecha_inicial
        session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = f"{fecha_inicial},{fecha_inicial}"
        
        # Ingresar la fecha final
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").setFocus()
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").caretPosition = 0
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = fecha_final
        session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = f"{fecha_final},{fecha_final}"
        
        # Ejecutar la transacción COOISPI
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        print("Transacción COOISPI ejecutada correctamente.")

        # --- Exportación del archivo ---
        # Eliminar el archivo previo (si existe)
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
            print("Archivo existente eliminado.")
        
        # Abrir el menú de exportación (IDs obtenidos del Scripting Tracker)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
        session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")
        
        # Configurar la ruta y el nombre del archivo en la ventana de exportación
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FOLDER_PATH
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Backup_COOISPI.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("Backup_COOISPI.xlsx")
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(5)  # Esperar a que se complete la exportación
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)
        
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        
        
        # --- Cierre de Excel ---
        cerrar_excel()
        print(f"Archivo exportado correctamente a: {EXCEL_FILE}")
        
    except Exception as e:
        print(f"Error al ejecutar la exportación o cerrar la transacción: {e}")

def main():
    session = conectar_sap()
    if session:
        ejecutar_exp(session)
        print("Proceso completado.")

if __name__ == "__main__":
    main()

        
        
