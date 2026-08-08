import json
import time
import subprocess
import os
import pandas as pd
import pyautogui
import win32gui, win32con

# Configuración de archivos de entrada (Plantillas por defecto)
ARCHIVO_COORDENADAS = "coordenadas_template.json"
ARCHIVO_EXCEL = "liquidacion_template.xlsx"
RUTA_INMOSOFT_DEFAULT = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Inmosoft\Inmosoft.lnk"

def formatear_monto_inmosoft(val):
    if pd.isna(val):
        return "0,00"
    s_val = str(val).replace('$', '').strip()
    if ',' in s_val and '.' in s_val:
        s_val = s_val.replace('.', '')
    s_val = s_val.replace('.', ',')
    try:
        val_float = float(s_val.replace(',', '.'))
        return f"{val_float:.2f}".replace('.', ',')
    except ValueError:
        return s_val

def enfocar_inmosoft():
    try:
        hwnd = win32gui.FindWindow(None, "Inmosoft")
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        else:
            if os.path.exists(RUTA_INMOSOFT_DEFAULT):
                subprocess.Popen(RUTA_INMOSOFT_DEFAULT)
                print("[+] Abriendo Inmosoft... Esperando 15s para inicio/login.")
                time.sleep(15)
            else:
                print("[!] No se encontró el acceso directo por defecto de Inmosoft.")
                print("    Por favor, abrí la aplicación manualmente antes de ejecutar el script.")
    except Exception as e:
        print(f"[!] Error enfocando Inmosoft: {e}")

def cargar_conceptos():
    if not os.path.exists(ARCHIVO_COORDENADAS):
        print(f"❌ ERROR: No se encontró el archivo '{ARCHIVO_COORDENADAS}'. Ejecutá primero 'calibrar_coordenadas.py'.")
        return

    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"❌ ERROR: No existe el archivo '{ARCHIVO_EXCEL}'. Verificá el nombre de la planilla.")
        return

    with open(ARCHIVO_COORDENADAS, "r", encoding="utf-8") as f:
        coords = json.load(f)

    df = pd.read_excel(ARCHIVO_EXCEL)
    df['Cód. Inmueble'] = df['Cód. Inmueble'].ffill()
    df = df.dropna(subset=['Concepto', 'Monto'])

    unidades = df.groupby('Cód. Inmueble', sort=False)

    enfocar_inmosoft()
    pyautogui.PAUSE = 0.4

    for cod_inmueble, grupo in unidades:
        print(f"\n[+] Procesando Inmueble: {cod_inmueble} ({len(grupo)} conceptos)")

        # 01 - Buscador rápido
        pyautogui.click(coords["01_paso_4_busqueda"])
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.write(str(cod_inmueble).strip(), interval=0.05)
        pyautogui.press('enter')
        time.sleep(1.2)

        # 02 - Opciones de contrato
        pyautogui.click(coords["02_paso_5_opciones"])
        time.sleep(0.4)

        # 03 - Programar conceptos para inquilino
        pyautogui.click(coords["03_paso_6_programar"])
        time.sleep(0.8)

        # Loop de conceptos por unidad
        for _, fila in grupo.iterrows():
            concepto = str(fila['Concepto']).strip()
            monto_formateado = formatear_monto_inmosoft(fila['Monto'])

            print(f"   -> {concepto} | ${monto_formateado}")

            # 04 - Agregar
            pyautogui.click(coords["04_paso_7_agregar"])
            time.sleep(0.5)

            # 05 - Nombre Concepto
            pyautogui.click(coords["05_paso_8_concepto"])
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(concepto, interval=0.05)

            # 06 - Monto
            pyautogui.click(coords["06_paso_9_monto"])
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(monto_formateado, interval=0.05)

            # 08 - Aceptar
            pyautogui.click(coords["08_paso_11_aceptar"])
            time.sleep(0.5)

        # 09 - Aplicar Cambios
        pyautogui.click(coords["09_paso_13_aplicar_cambios"])
        time.sleep(0.8)

        # Reset a Dashboard
        pyautogui.press('esc')
        time.sleep(0.3)
        pyautogui.press('esc')
        time.sleep(0.5)

    print("\n[✓] Carga de todas las unidades finalizada.")

if __name__ == "__main__":
    cargar_conceptos()
