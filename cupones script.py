import json
import time
import os
import pandas as pd
import pyautogui
import pyperclip
import win32gui, win32con

# Configuración de rutas y archivos de entrada (Plantillas por defecto)
CARPETA_DESTINO = os.path.abspath("CUPONES_DESCARGADOS")
ARCHIVO_COORDENADAS = "coordenadas_template.json"
ARCHIVO_EXCEL = "liquidacion_template.xlsx"

if not os.path.exists(CARPETA_DESTINO):
    os.makedirs(CARPETA_DESTINO)

def enfocar_inmosoft():
    try:
        hwnd = win32gui.FindWindow(None, "Inmosoft")
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
        else:
            print("[!] Inmosoft no está abierto. Abrí la aplicación manualmente antes de continuar.")
    except Exception as e:
        print(f"[!] Error enfocando Inmosoft: {e}")

def descargar_cupones():
    # 1. Validaciones previas de entorno
    if not os.path.exists(ARCHIVO_COORDENADAS):
        print(f"❌ ERROR: No se encontró '{ARCHIVO_COORDENADAS}'. Ejecutá primero 'calibrar_coordenadas.py'.")
        return

    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"❌ ERROR: No existe el archivo '{ARCHIVO_EXCEL}'. Verificá el nombre de la planilla.")
        return

    try:
        with open(ARCHIVO_COORDENADAS, "r", encoding="utf-8") as f:
            coords = json.load(f)
    except Exception as e:
        print(f"❌ ERROR al leer el JSON de coordenadas: {e}")
        return

    try:
        df = pd.read_excel(ARCHIVO_EXCEL)
        df['Cód. Inmueble'] = df['Cód. Inmueble'].ffill()
        inmuebles = df['Cód. Inmueble'].dropna().unique()
    except Exception as e:
        print(f"❌ ERROR al procesar la planilla Excel: {e}")
        return

    enfocar_inmosoft()
    pyautogui.PAUSE = 0.5

    for idx, cod_inmueble in enumerate(inmuebles, 1):
        cod_clean = str(cod_inmueble).strip()
        nombre_pdf = f"Cupon_{cod_clean}.pdf"
        ruta_completa_pdf = os.path.join(CARPETA_DESTINO, nombre_pdf)

        print(f"[{idx}/{len(inmuebles)}] Descargando: {cod_clean} -> {nombre_pdf}")

        # 1. Buscador rápido
        pyautogui.click(coords["01_busqueda"])
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'a')
        
        pyperclip.copy(cod_clean)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1.8)

        # 2. Útiles / Opciones
        pyautogui.click(coords["02_utiles_opciones"])
        time.sleep(0.5)

        # 3. Descargar cupón como PDF
        pyautogui.click(coords["03_descargar_pdf_directo"])
        time.sleep(1.8)

        # 4. Cuadro Guardar como de Windows
        pyperclip.copy(ruta_completa_pdf)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.4)
        pyautogui.press('enter')  # Confirma guardado
        time.sleep(1.2)

        # 5. Cierra el cartel "¿Desea abrir el PDF?"
        pyautogui.press('esc')
        time.sleep(0.4)

        # 6. Cierra la ficha del contrato y vuelve al Dashboard
        pyautogui.press('esc')
        time.sleep(0.8)

    print(f"\n[✓] Descarga masiva finalizada. Guardados en: {CARPETA_DESTINO}")

if __name__ == "__main__":
    descargar_cupones()