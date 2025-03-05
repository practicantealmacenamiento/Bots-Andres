import os
import time
import openpyxl
import win32com.client
import pandas as pd
from Generales import send_email, close_excel

# ────────────────────────
# Configuración inicial  
# ────────────────────────
SEND_EMAILS = True
AFK = True

# ────────────────────────
# Configuración de Rutas
# ────────────────────────
FOLDER_PATH = r"C:\TEMP\Backup Existencias"
LX03_FILE = os.path.join(FOLDER_PATH, "LX03.xlsx")
IH09_FILE = os.path.join(FOLDER_PATH, "IH09.xlsx")
LX03_ACTUALIZADO_FILE = os.path.join(FOLDER_PATH, "LX03_actualizado.xlsx")

# ────────────────────────
# Funciones de SAP
# ────────────────────────
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

def ejecutar_LX03(session):
    """
    Ejecuta la transacción LX03 en SAP utilizando los IDs proporcionados y exporta el reporte a LX03.xlsx.
    """
    try:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00019")
        session.findById("wnd[0]/usr/chkPMITB").selected = True
        session.findById("wnd[0]/usr/ctxtS1_LGNUM").text = "pro"
        session.findById("wnd[0]/usr/chkPMITB").setFocus()
        session.findById("wnd[0]/usr/btn%_S1_LGTYP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[6]").press()
        chk30 = session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/30[0,30]/chk[1,30]")
        chk30.selected = True
        chk30.setFocus()
        wnd2_usr = session.findById("wnd[2]/usr")
        wnd2_usr.verticalScrollbar.position = 1
        wnd2_usr.verticalScrollbar.position = 2
        wnd2_usr.verticalScrollbar.position = 3
        wnd2_usr.verticalScrollbar.position = 4
        chk31 = session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/31[0,31]/chk[1,31]")
        chk31.selected = True
        chk31.setFocus()
        wnd2_usr.verticalScrollbar.position = 5
        wnd2_usr.verticalScrollbar.position = 6
        wnd2_usr.verticalScrollbar.position = 7
        wnd2_usr.verticalScrollbar.position = 8
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/30[0,30]/chk[1,30]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/31[0,31]/chk[1,31]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/33[0,33]/chk[1,33]").selected = True
        chk34 = session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/34[0,34]/chk[1,34]")
        chk34.selected = True
        chk34.setFocus()
        wnd2_usr.verticalScrollbar.position = 9
        wnd2_usr.verticalScrollbar.position = 10
        wnd2_usr.verticalScrollbar.position = 11
        wnd2_usr.verticalScrollbar.position = 12
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/31[0,31]/chk[1,31]").selected = True
        chk33 = session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/33[0,33]/chk[1,33]")
        chk33.selected = True
        chk33.setFocus()
        wnd2_usr.verticalScrollbar.position = 13
        wnd2_usr.verticalScrollbar.position = 14
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/32[0,32]/chk[1,32]").selected = True
        chk34 = session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/34[0,34]/chk[1,34]")
        chk34.selected = True
        chk34.setFocus()
        wnd2_usr.verticalScrollbar.position = 15
        wnd2_usr.verticalScrollbar.position = 16
        wnd2_usr.verticalScrollbar.position = 17
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/32[0,32]/chk[1,32]").selected = True
        chk33 = session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/33[0,33]/chk[1,33]")
        chk33.selected = True
        chk33.setFocus()
        wnd2_usr.verticalScrollbar.position = 34
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/17[0,17]/chk[1,17]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/18[0,18]/chk[1,18]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/19[0,19]/chk[1,19]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/23[0,23]/chk[1,23]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/24[0,24]/chk[1,24]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/25[0,25]/chk[1,25]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/26[0,26]/chk[1,26]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/27[0,27]/chk[1,27]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/28[0,28]/chk[1,28]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/29[0,29]/chk[1,29]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/30[0,30]/chk[1,30]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/31[0,31]/chk[1,31]").selected = True
        session.findById("wnd[2]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/31[0,31]/chk[1,31]").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\TEMP\Backup Existencias"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "LX03.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        print("Transacción LX03 ejecutada y exportada a:", LX03_FILE)
    except Exception as e:
        print("Error durante la ejecución de LX03:", e)
        
        
def email():
    subject = "Permanencias Milagros"
    emails = [
        "ferney.correa@prebel.com.co",
        "practicante.almacenamiento@prebel.com.co"
    ]
    copy_to = []
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
            <p>Adjunto el informe de permanencias Milagros.</p>
            <p>¡Feliz día!</p>
            <br>
        </body>
        </html>
    """
    
    attachments = [LX03_ACTUALIZADO_FILE]
    send_email(
        subject=subject,
        emails="; ".join(emails),
        emails_cc="; ".join(copy_to),
        files_Attachment=attachments,
        html_contend=body
    )

def close_excel(tiempo_espera=2):
    try:
        print(f"Esperando {tiempo_espera} segundos antes de cerrar Excel...")
        time.sleep(tiempo_espera)
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            workbook = excel.Workbooks(os.path.basename(LX03_FILE))
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
        
# ────────────────────────
# Procesamiento de datos 
# ────────────────────────

def procesar_excel():
    """
    Realiza un merge entre los archivos LX03.xlsx e IH09.xlsx.
    Agrega la columna 'Texto breve de material' del archivo IH09.xlsx 
    al archivo LX03.xlsx usando como base la columna común 'Material'.
    """
    try: 
        # Lee los archivos 
        df_LX03 = pd.read_excel(LX03_FILE)
        df_IH09 = pd.read_excel(IH09_FILE)
        
        print("Columnas LX03", df_LX03.columns)
        print("Columnas IH09", df_IH09.columns)
        
        # Asegurarse de que la columna "Material" sea de tipo string, sin espacios extra y en mayúsculas
        df_LX03['Material'] = df_LX03['Material'].astype(str).str.strip().str.upper()
        df_IH09['Material'] = df_IH09['Material'].astype(str).str.strip().str.upper()
        
        # Filtrar los materiales que comienzan con "MP"
        df_LX03 = df_LX03[df_LX03['Material'].str.startswith("MP")]
        df_IH09 = df_IH09[df_IH09['Material'].str.startswith("MP")]
        
        # Realiza el merge para traer 'Texto breve de material'
        df_resultado = df_LX03.merge(
            df_IH09[["Material", "Texto breve de material"]],
            on="Material",
            how="left"
        )
        
        # Reordenar columnas: insertar 'Texto breve de material' justo después de 'Material'
        cols = list(df_resultado.columns)
        if "Texto breve de material" in cols:
            cols.remove("Texto breve de material")
            material_index = cols.index("Material")
            cols.insert(material_index + 1, "Texto breve de material")
            df_resultado = df_resultado[cols]
        
        # Guardar el resultado en un nuevo archivo 
        df_resultado.to_excel(LX03_ACTUALIZADO_FILE, index=False)
        print("Archivo actualizado generado", LX03_ACTUALIZADO_FILE)
    except Exception as e:
        print("Error en el procesamieno de Excel:", e)
        

def main():
    session = conectar_sap()
    if session is None:
        return
    
    ejecutar_LX03(session)
    
    procesar_excel()
        
    close_excel(tiempo_espera=2)
    
    if SEND_EMAILS:
        email()

if __name__ == "__main__": 
    main()

# Fin del script
