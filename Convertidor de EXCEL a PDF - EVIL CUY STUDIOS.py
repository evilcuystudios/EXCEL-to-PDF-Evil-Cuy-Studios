import os
import win32com.client
import time
import tkinter as tk
from tkinter import filedialog
import subprocess

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    # El usuario elige la carpeta de origen
    folder_selected = filedialog.askdirectory(title="EVIL CUY STUDIOS - SELECCIONA CARPETA DE EXCEL")
    root.destroy()
    return folder_selected

def print_hacker_cuy():
    print(r"""
    ____________________________________________________________

       _     _          
      (c).-.(c)         [ EVIL CUY STUDIOS ]
       / . . \          --------------------
      =\__Y__/=         USER: ACCOUNTANT_CUY 
      /       \         TASK: EXCEL_TO_PDF
      \_______/         STATUS : INFILTRATING...
       `"" ""`          AUTHOR : PROCYON
    ____________________________________________________________
    """)

def convert_excel():
    clear_screen()
    print_hacker_cuy()
    
    print(" [>] Esperando selección de carpeta...")
    path = select_folder()

    if not path:
        print("\n [!] OPERACIÓN CANCELADA: No seleccionaste ninguna carpeta.")
        time.sleep(2)
        return

    # Normalizamos la ruta para evitar conflictos de Windows
    path = os.path.abspath(path)
    os.chdir(path)
    
    files = [f for f in os.listdir(path) if f.lower().endswith(('.xlsx', '.xls', '.csv'))]

    if not files:
        print(f"\n [!] ALERTA: No hay archivos Excel en: {path}")
        input("\n PRESIONA [ENTER] PARA SALIR...")
        return

    # LA BÓVEDA SE CREA JUSTO AQUÍ (En la carpeta seleccionada)
    output_folder = os.path.join(path, "RESULTADOS_PDF")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    print(f" [+] CARPETA ORIGEN: {path}")
    print(f" [+] OBJETIVOS LOCALIZADOS: {len(files)}")
    print(f" [+] BÓVEDA LOCAL: {output_folder}")
    print(" ------------------------------------------------------------")

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        for index, file in enumerate(files, 1):
            input_path = os.path.join(path, file)
            output_name = os.path.splitext(file)[0] + ".pdf"
            output_path = os.path.join(output_folder, output_name)
            
            print(f" > [{index}/{len(files)}] PROCESANDO: {file[:25]}...", end="\r")
            
            wb = excel.Workbooks.Open(input_path)
            # 0 representa el formato PDF en la API de Excel
            wb.ExportAsFixedFormat(0, output_path)
            wb.Close(False)
            
            print(f" > [{index}/{len(files)}] GUARDADO EN BÓVEDA: {output_name}            ")
            
        print("\n [+] ABRIENDO BÓVEDA LOCAL...")
        # Abrimos la carpeta específica de resultados
        subprocess.Popen(f'explorer "{output_folder}"')
            
    except Exception as e:
        print(f"\n\n [X] ERROR EN LA OPERACIÓN: {e}")
    finally:
        if excel:
            excel.Quit()
            del excel
        
        print("\n ------------------------------------------------------------")
        print(f" >> OPERACIÓN FINALIZADA CON ÉXITO.")
        print(" ------------------------------------------------------------")
        input(" PRESIONA [ENTER] PARA CERRAR...")

if __name__ == "__main__":
    convert_excel()