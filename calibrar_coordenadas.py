import json
import pyautogui

# Puntos para el flujo de Carga Masiva de Conceptos
PUNTOS_CARGA = [
    ("01_paso_4_busqueda", "Buscador Rápido de Contratos (Dashboard)"),
    ("02_paso_5_opciones", "Botón 'Opciones de contrato'"),
    ("03_paso_6_programar", "Opción 'Programar conceptos para inquilino'"),
    ("04_paso_7_agregar", "Botón 'Agregar' (Ventana Conceptos Programados)"),
    ("05_paso_8_concepto", "Campo 'Concepto' (Ventana Nuevo Concepto)"),
    ("06_paso_9_monto", "Campo 'Monto'"),
    ("07_paso_10_aplica_en", "Campo/Desplegable 'Aplica en'"),
    ("08_paso_11_aceptar", "Botón 'Aceptar' (Ventana Nuevo Concepto)"),
    ("09_paso_13_aplicar_cambios", "Botón 'Aplicar cambios' (Ventana Conceptos Programados)")
]

# Puntos para el flujo de Descarga Directa de Cupones PDF
PUNTOS_CUPONES = [
    ("01_busqueda", "Buscador Rápido de Contratos (Dashboard)"),
    ("02_utiles_opciones", "Desplegable Útiles / Opciones de Contrato"),
    ("03_descargar_pdf_directo", "Opción 'Descargar cupón de pago como pdf'"),
    ("04_confirmar_guardar", "Botón 'Guardar' (Ventana emergente de Windows)")
]

ARCHIVO_SALIDA = "coordenadas_template.json"

def cargar_coordenadas_existentes():
    try:
        with open(ARCHIVO_SALIDA, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def ejecutar_mapeo(lista_puntos, coordenadas_dict):
    for id_punto, descripcion in lista_puntos:
        input(f"👉 [{id_punto}] Posicioná el mouse sobre [{descripcion}] y presioná ENTER...")
        x, y = pyautogui.position()
        coordenadas_dict[id_punto] = [x, y]
        print(f"   ✓ Capturado [{id_punto}]: X={x}, Y={y}\n")

def calibrar():
    print("=======================================================")
    print("   CALIBRADOR UNIFICADO DE COORDENADAS - INMOSOFT RPA ")
    print("=======================================================")
    print("Seleccioná la opción que deseas calibrar:")
    print("1. Solo Carga de Conceptos (9 Puntos)")
    print("2. Solo Descarga de Cupones PDF (4 Puntos)")
    print("3. CALIBRACIÓN COMPLETA (Ambos flujos)")
    print("=======================================================")
    
    opcion = input("Ingresá opción (1, 2 o 3): ").strip()
    coordenadas = cargar_coordenadas_existentes()

    if opcion == "1":
        print("\n--- Calibrando Puntos de Carga de Conceptos ---")
        ejecutar_mapeo(PUNTOS_CARGA, coordenadas)
    elif opcion == "2":
        print("\n--- Calibrando Puntos de Descarga de Cupones ---")
        ejecutar_mapeo(PUNTOS_CUPONES, coordenadas)
    elif opcion == "3":
        print("\n--- Calibración Completa ---")
        ejecutar_mapeo(PUNTOS_CARGA, coordenadas)
        ejecutar_mapeo(PUNTOS_CUPONES, coordenadas)
    else:
        print("❌ Opción no válida. Cancelando.")
        return

    # Guardar en JSON consolidado
    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as f:
        json.dump(coordenadas, f, indent=4)

    print("=======================================================")
    print(f"✅ CALIBRACIÓN GUARDADA EXITOSAMENTE")
    print(f"Archivo actualizado: '{ARCHIVO_SALIDA}'")
    print("=======================================================")

if __name__ == "__main__":
    calibrar()
