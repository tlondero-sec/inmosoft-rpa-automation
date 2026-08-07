# 🤖 Inmosoft RPA Automation — Legacy Accounting & PDF Billing Engine

![Python](https://img.shields.io/badge/Language-Python%203.x-blue?style=for-the-badge&logo=python)
![Automation](https://img.shields.io/badge/Focus-RPA%20%2F%20Legacy%20Systems-orange?style=for-the-badge)
![OS](https://img.shields.io/badge/Platform-Windows%20x64-0078D6?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

## 📌 Resumen del Proyecto

Solución de **Automatización de Procesos Robóticos (RPA)** desarrollada en Python para optimizar la carga contable masiva de liquidaciones y la emisión/descarga directa de cupones de pago en PDF dentro de la plataforma **Inmosoft** (sistema empresarial ERP legacy de 32 bits).

El desarrollo resuelve de forma nativa cuellos de botella de la API de Windows, desbordamientos de búfer en cajas de texto de sistemas de 32 bits y saturación de memoria mediante la inyección directa por portapapeles y calibración por coordenadas absolutas.

---

## 🛠️ Desafíos Técnicos & Soluciones de Ingeniería

### 1. Desbordamiento de Búfer en Cargas Numéricas (Error -99999999,00)
* **Problema:** Inmosoft es una aplicación legacy de 32 bits que no procesa puntos decimales ni separadores de miles al recibir texto pegado o tiapeado. Al ingresar valores como `1.000,00`, el sistema sufre un integer overflow asignando `-99999999,00`.
* **Solución:** Se diseñó una función de sanitización estricta (`formatear_monto_inmosoft`) que limpia símbolos monetarios (`$`), elimina puntos separadores de miles y transforma obligatoriamente cualquier entrada numérica al formato con coma decimal de precisión fija (`1000,00`).

### 2. Saturación de Interrupciones GUI (Cuelgues WOW64)
* **Problema:** La simulación acelerada de pulsaciones de teclado mediante librerías de automatización satura la cola de eventos de Windows en la capa de compatibilidad WOW64 (32-bit sobre 64-bit), colapsando el software.
* **Solución:** Se implementó inyección atómica por portapapeles de Windows (`pyperclip` + `Ctrl+V`) acompañada de delays estratégicos (`pyautogui.PAUSE = 0.4s` a `0.5s`) y esperas adaptativas de lectura en base de datos.

### 3. Agrupamiento de Estructuras No Relacionales en Excel
* **Problema:** La planilla de liquidación contable agrupa los conceptos por bloque pero solo declara el `Cód. Inmueble` en la primera fila de la unidad funcional.
* **Solución:** Uso del algoritmo de propagación `df['Cód. Inmueble'].ffill()` de `pandas` para autocompletar la clave foránea hacia abajo, agrupando dinámicamente con `.groupby()` sin alterar el orden original.

### 4. Sobreescritura de Archivos en Descarga Masiva
* **Problema:** Inmosoft exporta reportes PDF con nombres estáticos por propietario. Cuando múltiples unidades pertenecen al mismo titular, el archivo local se sobreescribe.
* **Solución:** Intercepción de la ventana *"Guardar como"* de Windows (`win32gui`) e inyección directa de la ruta absoluta con nombrado único en tiempo real (`C:\...\CUPONES_DESCARGADOS\Cupon_CODIGO.pdf`).

---

## 🏗️ Arquitectura de Scripts

```text
inmosoft-rpa-automation/
│
├── 📜 calibrar_coordenadas.py      # Calibrador modular interactivo para mapeo de pantalla
├── 📜 validar_excel.py             # Script de sanity check / pre-flight validation de liquidacion.xlsx
├── 📜 inmosoft_carga_conceptos.py  # Bot de carga masiva e inyección de conceptos a Inmosoft
├── 📜 descargar_cupones_estable.py # Bot de descarga masiva de cupones PDF con nombrado dinámico
│
├── 📂 CUPONES_DESCARGADOS/         # Directorio de salida de reportes PDF generados
├── 📊 liquidacion.xlsx             # Planilla de origen de datos (Cód. Inmueble, Concepto, Monto)
└── ⚙️ coordenadas_monitor_teo.json # Mapeo de coordenadas en JSON para independencia de resolución
