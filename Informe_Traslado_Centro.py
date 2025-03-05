from datetime import datetime
import os
import time
import sapgui
import win32com.client
from Generales import close_excel, send_email

# Configuración inicial
SEND_EMAILS = True
AFK = True

# Configuración de destino
FOLDER_PATH = r"C:\TEMP\Informe_Traslados_ZMM78"
EXCEL_FILE = os.path.join(FOLDER_PATH, "Informe_ZMM78.xlsx")

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
    
def close_excel(tiempo_espera=1):
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
        workbook = excel.Workbooks("Informe_ZMM78.xlsx")
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



def esperar_descarga(archivo, timeout=30):
    """
    Espera hasta que el archivo se descargue y tenga la fecha de modificación actual.
    timeout: tiempo máximo de espera en segundos.
    """
    print("Esperando a que el informe se descargue...")
    elapsed_time = 0
    while elapsed_time < timeout:
        if os.path.exists(archivo):
            last_modified_time = os.path.getmtime(archivo)
            last_modified_date = datetime.fromtimestamp(last_modified_time).strftime("%Y%m%d")
            today_date = datetime.now().strftime("%Y%m%d")

            if last_modified_date == today_date:
                print(f"El archivo {archivo} ha sido actualizado correctamente.")
                return True
        time.sleep(1)
        elapsed_time += 1
    print(f"Error: El archivo {archivo} no se actualizó dentro del tiempo límite.")
    return False


def ejecutar_zmm78(session):
    """Ejecuta la transacción ZMM78"""
    try:
        fecha_actual = datetime.now().strftime("%Y%m%d")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmm78"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btnBTN8").press()
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = f"{fecha_actual},{fecha_actual}"
        session.findById("wnd[0]/usr/chkP_SHOW").selected = True
        session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = "3000"
        session.findById("wnd[0]/usr/ctxtSO_LGORT-LOW").text = ""
        session.findById("wnd[0]/usr/chkP_SHOW").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").currentCellRow = 7
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "7"
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FOLDER_PATH
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Informe_ZMM78.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        
        print(f"Archivo exportado correctamente a: {EXCEL_FILE}")
        
    except Exception as e:
        print(f"Error al ejecutar zmm78 o exportar: {e}")


def email():
        subject = "Traslado Centro"
        emails = [
            # "practicante.almacenamiento@prebel.com.co",
            "lideres.recibopt@prebel.com.co",
            "recibo.medellin@prebel.com.co",
            "recibo.disnal@prebel.com.co",
        ]
        copy_to = [
            "practicante.almacanemiento@prebel.com.co"
            "yakelin.rojas@prebel.com.co",
            "juan.espinosa@prebel.com.co",
            "carlos.rivas@prebel.com.co",
            "ferney.correa@prebel.com.co",
            "inventarios.medellin@prebel.com.co",
        ]
        body = """
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Informe ZMM78</title>
            </head>
            <body>
                <h2>Buenos dias</h2>
                <p>Adjunto el informe traslado centro 1000/3000.</p>
                <p>¡Que tengan un feliz día!</p>
                <br>
            </body>
            </html>
        """
        attachments = [EXCEL_FILE]
        send_email(
            subject=subject,
            emails="; ".join(emails),
            emails_cc="; ".join(copy_to),
            files_Attachment=attachments,
            html_contend=body
        )

# Ejecución del script
session = conectar_sap()

if session:
    ejecutar_zmm78(session)
    
    # 🔹 Agregar un `time.sleep()` de seguridad
    time.sleep(5)
    
    close_excel()
    
    # 🔹 Esperar hasta que el archivo realmente se descargue
    if esperar_descarga(EXCEL_FILE):
        if SEND_EMAILS:
            email()
    else:
        print("No se enviará el correo porque el informe no se generó correctamente.")


        
        