
# 🤖 Inmosoft RPA Automation — Legacy Accounting & PDF Billing Engine

![Python](https://img.shields.io/badge/Language-Python%203.x-blue?style=for-the-badge&logo=python)
![Automation](https://img.shields.io/badge/Focus-RPA%20%2F%20Legacy%20Systems-orange?style=for-the-badge)
![OS](https://img.shields.io/badge/Platform-Windows%20x64-0078D6?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

📌 Resumen del Proyecto

Script en Python para automatizar la carga de liquidaciones y la descarga de cupones PDF en el ERP Inmosoft (Windows 32-bit).

Estrategias y comportamiento actual:

    Inyección por portapapeles (Clipboard): Agiliza el pegado de texto en cajas de entrada (reduce errores de tipeo, aunque el ERP mantiene cierres inesperados por memory leaks en ejecuciones largas).

    Sanitización de montos: Previene el desbordamiento de búfer (-99999999,00) forzando el formato con coma decimal de precisión fija.

    Calibración por coordenadas: Mapea la interfaz gráfica mediante un JSON unificado de coordenadas absolutas para interactuar con la GUI sin depender de controles internos.

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

```
Control Actual (Supervisión Manual Asistida)

👁️ El bot opera a ciegas mediante coordenadas fijas (Blind Automation), por lo que requiere supervisión constante por parte del operador:

    Pausas estratégicas (0.5s a 1.8s): Los delays entre acciones no garantizan la estabilidad del ERP; están puestos adrede para darle tiempo al software de procesar y, principalmente, para darle al operador una ventana de reacción donde presionar Esc o frenar la ejecución si la GUI se congela o descoordina.

    Desfase Terminal vs. GUI: La consola imprime la acción programada antes de enviarla. Si el programa se frena o pierde el foco, la terminal seguirá mostrando pasos que no están ocurriendo en pantalla, ya que el script no valida el estado real de las ventanas.


## 🔄 Flujo Operativo Mapeado

```mermaid
graph TD
    A[Inicio: liquidacion.xlsx] --> B[validar_excel.py: Sanity Check]
    B -->|Éxito| C[ffill & GroupBy por Cód. Inmueble]
    C --> D[Enfocar Inmosoft con win32gui]
    
    subgraph Fase 1: Carga de Conceptos
        D --> E[Buscador Rápido: Cód. Inmueble]
        E --> F[Programar Conceptos para Inquilino]
        F --> G[Inyección de Concepto y Monto Sanitizado]
        G --> H[Aplicar Cambios & Reset con Esc]
    end

    subgraph Fase 2: Emisión y Descarga de PDF
        H --> I[Navegación: Útiles -> Descargar cupón como PDF]
        I --> J[Inyección de Ruta Absoluta en Windows Save As]
        J --> K[Cierre Limpio de Popups & Retorno a Dashboard]
    end
    
    K --> L[Fin: Cupones Exportados en CUPONES_DESCARGADOS]

```

---

## 💻 Código Destacado: Motor de Sanitización y Manejo de Portapapeles

### Sanitización de Montos para ERP 32-bit:

```python
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

```

### Inyección Atómica de Rutas Absolutas (Evitando diálogo de sobreescritura):

```python
# Inyección directa por portapapeles en la API de Windows
pyperclip.copy(ruta_completa_pdf)
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.2)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.4)
pyautogui.press('enter')  # Confirma guardado

```

---

## 📋 Pre-requisitos e Instalación

> [!IMPORTANT]
> **🔴 REQUISITO OBLIGATORIO DE ENTORNO: POWERSHELL 7+ (ADMINISTRADOR)**
> 
> Debido a la interacción directa con las APIs de Windows (`win32gui` / `win32con`) y a la simulación de eventos de teclado/mouse a bajo nivel, **los scripts NO funcionarán correctamente en la consola clásica de Windows Command Prompt (cmd.exe) ni en la versión legacy de Windows PowerShell 5.1**.
> 
> * **Motor Requerido:** **PowerShell 7.x+** (probado y validado en **PowerShell 7.4+ / 7.6.x Core**).
> * **Elevación de Privilegios:** La terminal DEBE ejecutarse obligatoriamente con el rol de **Administrador** (`Run as Administrator`). De lo contrario, las llamadas a `SetForegroundWindow` e inyección de portapapeles sobre la ventana de Inmosoft serán bloqueadas por las políticas de seguridad (UAC / UIPI) de Windows.

# 1. Abrir PowerShell 7 (pwsh.exe) como Administrador
# 2. Navegar a la carpeta del proyecto
cd "C:\RPA\inmosoft-rpa-automation"

# 3. Validar versión del entorno (debe ser 7.x o superior)
$PSVersionTable.PSVersion

# 4. Ejecutar el validador y los bots
python validar_excel.py
python inmosoft_carga_conceptos.py

1. **Entorno Python 3.8+** (Windows 10/11 x64).
2. **Instalación de Dependencias:**
```bash
pip install pandas openpyxl pyautogui pyperclip pywin32

```


3. **Ejecución de Pre-flight Check:**
```bash
python validar_excel.py

```


4. **Calibración de Pantalla (Solo si cambia de monitor o resolución):**
```bash
python calibrar_coordenadas.py

```


5. **Ejecución de Automatismos:**
```bash
python inmosoft_carga_conceptos.py
python descargar_cupones_estable.py

```
### 📊 Estructura del Archivo de Entrada (`liquidacion_template.xlsx`)

El archivo de Excel debe respetar la estructura de bloques agrupados por unidad funcional. No es necesario repetir el código de inmueble en cada fila; el bot aplica *forward fill* automáticamente.

![Plantilla de Excel](./img/00-template-excel.png)

> **Nota sobre el formato:** Asegurate de mantener los encabezados exactos `Cód. Inmueble`, `Concepto` y `Monto`.


---
#### 🚀 Roadmap Técnico de Mejora (Context-Aware RPA):
Para evolucionar de una automatización asistida a una ejecución autónoma y resiliente, se contemplan las siguientes vías de refactorización:
1. **Verificación de Estado por Visión Computacional (OpenCV / PyScreeze):** Implementar reconocimiento de patrones visuales para confirmar la presencia de botones ("Agregar", "Aceptar") e íconos de carga antes de enviar clics.
2. **Inspección Dinámica de Elementos (Win32 API / Pywinauto):** Sustituir coordenadas absolutas por la lectura directa de *handles* de ventanas de Windows (`HWND`), detectando cambios de título o controles habilitados.
3. **Manejo de Excepciones por Timeout:** Reemplazar los `sleep()` fijos por un patrón *Wait Until Visible* con tiempo de espera máximo.
---

## 🛠️ Desafíos Técnicos & Soluciones de Ingeniería

### 🐛 Evidencia de Bug Legacy: Desbordamiento de Búfer (-99999999,00)

El motor gráfico de Inmosoft (32-bit) sufre una falla de parseo numérico cuando se envían caracteres o formatos no sanitizados en la caja de texto `Monto`. Esto causa un integer overflow asignando el valor límite por defecto `-99999999,00`.

| 1. Entrada de Texto | 2. Integer Overflow resultante |
| :---: | :---: |
| ![Monto Ingresado](./img/01-buffer-overflow-input.png) | ![Resultado Corrupto](./img/02-buffer-overflow-result.png) |
| *Intento de carga de monto.* | *Resultado desbordado en la GUI (`-99999999,00`).* |

> 🛡️ **Solución implementada en el Bot:** La función `formatear_monto_inmosoft()` intercepta el dato del Excel y fuerza la conversión limpia a coma decimal de precisión fija previa a la Inyección por Portapapeles, previniendo la corrupción de montos en el ERP.
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

> [!WARNING]
> **⚠️ ESTADO DEL PROYECTO & DISCLAIMER DE USO EN PRODUCCIÓN**
> 
> Este script fue desarrollado como una **Prueba de Concepto (PoC) / Automatización Ad-Hoc** para operar sobre una interfaz GUI legacy no pensada para automatización. Debido a la naturaleza no determinista de la simulación de eventos de mouse/teclado y la inestabilidad propia del software base, el repositorio se comparte exclusivamente con fines educativos y de portafolio. **No se recomienda su ejecución desatendida en entornos de producción.**

---

### 🐛 Known Issues & Deuda Técnica (Bugs Conocidos)

Actualmente existen comportamientos anómalos detectados en el pipeline que se encuentran bajo investigación para futuras versiones:

1. **Colisión y Sobreescritura Sporádica de Cupones PDF:**
   * **Comportamiento:** A pesar del renombrado dinámico por ruta absoluta en la ventana *"Guardar como"*, si la API de Windows demora en renderizar el cuadro de diálogo por sobrecarga de I/O, el bot envía la secuencia de guardado antes de limpiar el campo de texto, provocando la sobreescritura con el nombre estático por defecto del ERP.
2. **Crashes Aleatorios de Inmosoft (Memory Leaks / Fault Tolerant):**
   * **Comportamiento:** El motor interno de Inmosoft (32 bits) presenta fugas de memoria al procesar ciclos largos de lectura/escritura de contratos. Tras 20-30 iteraciones seguidas, la aplicación lanza una excepción no controlada (`Access Violation / Stopped Working`) cerrándose de forma imprevista.
3. **Salteo Inexplorado de Unidades Funcionales (Omission Anomaly):**
   * **Comportamiento:** Durante la lectura e inyección de `liquidacion.xlsx`, en ejecuciones extensas el flujo saltea ocasionalmente 1 o 2 unidades consecutivas sin arrojar error sintáctico en consola. Se sospecha una desincronización de tiempos (*race condition*) entre el foco del cuadro de búsqueda y la respuesta de la base de datos local.

> 🛠️ **Contribuciones / Workarounds:** Si tenés sugerencias para mitigar la condición de carrera vía `win32api` o mediante *hooks* directos a la ventana de Windows, los Pull Requests son más que bienvenidos.
---

## 👤 Autor

**Teo Quimey Waldemar Londero**

*Analista de Ciberseguridad & Desarrollador RPA*

* [GitHub Profile](https://www.google.com/search?q=https://github.com/tlondero-sec)
* [Portfolio General](https://github.com/tlondero-sec/portfolio)

---


```
