# Inmosoft RPA — Carga de Liquidaciones y Descarga de Cupones

![Language](https://img.shields.io/badge/LANGUAGE-PYTHON_3.X-blue) ![Platform](https://img.shields.io/badge/PLATFORM-WINDOWS_X64-blue) ![License](https://img.shields.io/badge/LICENSE-MIT-green)

Conjunto de scripts en Python para automatizar la carga de conceptos contables y la descarga de cupones PDF en el ERP Inmosoft (Windows 32-bit).

## 🎥 Demostración
https://github.com/user-attachments/assets/1f994c49-d9ed-439f-88a0-aa8b8dfa3043
> *Ejecución del script experimental inyectando las liquidaciones directamente en los campos del ERP.*

---

## 1. Objetivo del Proyecto
* **Inyección por portapapeles (`Clipboard`):** Pega los textos directamente en los campos del ERP mediante la API del portapapeles en lugar de simular tipeo tecla por tecla. Agiliza la entrada y evita errores de caracteres.
* **Sanitización de montos:** Fuerza la conversión de los números al formato estricto con coma decimal (`1000,00`), evitando que la interfaz de Inmosoft rompa los valores al recibir puntos o formatos no parseados.
* **Mapeo por coordenadas:** Utiliza un archivo JSON de coordenadas absolutas (`X, Y`) para simular los clics sobre la interfaz gráfica sin depender de controles internos.

---

## 2. Pre-requisitos e Instalación

> [!IMPORTANT]
> **🔴 REQUISITO OBLIGATORIO DE ENTORNO: POWERSHELL 7+ (ADMINISTRADOR)**
> 
> Por la interacción directa con las APIs de Windows (`win32gui` / `win32con`) y la simulación de eventos de teclado/mouse a bajo nivel, **los scripts NO funcionarán como es debido en la consola clásica de Windows Command Prompt (cmd.exe) ni en la versión legacy de Windows PowerShell 5.1**.
> 
> * **Motor Requerido:** **PowerShell 7.x+** (validado en **PowerShell 7.4+ / 7.6.x Core**).
> * **Elevación de Privilegios:** La terminal DEBE ejecutarse con el rol de **Administrador** (`Run as Administrator`). Si no, las llamadas a `SetForegroundWindow` e inyección de portapapeles sobre la ventana de Inmosoft serán bloqueadas por las políticas de seguridad (UAC / UIPI) de Windows.

1. **Entorno Python 3.8+** (Windows 10/11 x64).
2. **Instalar dependencias:**
   ```powershell
   pip install pandas openpyxl pyautogui pyperclip pywin32
---

## 3. Estructura del Proyecto

```text
inmosoft-rpa-automation/
├── calibrar_coordenadas.py      # Script unificado para mapeo de puntos en pantalla (genera coordenadas_template.json)
├── Carga Inmosoft.py            # Script de carga de expensas y conceptos en el ERP
├── cupones script.py            # Script para descarga masiva de cupones en PDF
├── liquidacion_template.xlsx    # Planilla modelo de entrada de datos
├── 00-template-excel.png        # Captura de referencia de la planilla
├── 01-buffer-overflow-input.png # Captura de evidencia de entrada de texto
├── 02-buffer-overflow-result.png# Captura del error de desbordamiento de búfer
└── README.md                    # Documentación técnica

```

---

## 4. Estructura del Archivo de Entrada (`liquidacion_template.xlsx`)

El archivo de Excel debe seguir la estructura de bloques agrupados por unidad funcional. No es necesario repetir el código de inmueble en cada fila; la función `ffill()` autocompleta la clave hacia abajo de forma automática.

| Cód. Inmueble | Descripcion Inmueble | Concepto | Monto | Complejo |
| :--- | :--- | :--- | :--- | :--- |
| **DE15** | EA1D | IMPUESTO INMOBILIARIO MUNICIPAL | 1000,00 | |
| | | ABONO ASCENSOR | 1000,00 | |
| | | MANTENIMIENTO INGRESO EDIFICIO | 1000,00 | |
| **DE16** | EA2D | IMPUESTO INMOBILIARIO MUNICIPAL | 1000,00 | |
| | | ABONO ASCENSOR | 1000,00 | |
| | | Seguro Contra incendio Edificio Entero | 1000,00 | |
| **DE17** | EA4C | IMPUESTO INMOBILIARIO MUNICIPAL | 1000,00 | |
| | | ABONO ASCENSOR | 1000,00 | |
| | | LIMPIEZA | 1000,00 | |

> **Nota sobre el formato:** Asegúrate de mantener los encabezados exactos `Cód. Inmueble`, `Concepto` y `Monto`.

---

## 5. Modo de Uso y Supervisión Manual

Los scripts requieren atención humana constante durante la ejecución. No validan si las ventanas abrieron correctamente ni si el ERP se colgó; simplemente mueven el mouse y envían teclas a ciegas.

```powershell
# 1. Abrir PowerShell 7 (pwsh.exe) como Administrador
# 2. Ir a la carpeta del proyecto
cd "C:\RPA\inmosoft-rpa-automation"

# 3. Calibrar coordenadas (solo si cambias de monitor o resolución)
python calibrar_coordenadas.py

# 4. Ejecutar la carga o descarga
python "Carga Inmosoft.py"
python "cupones script.py"

```

* **Delays intencionales (`0.5s` a `1.8s`):** Le dan tiempo al programa a responder y dejan una ventana de reacción para que el operador presione `Esc`. La tecla `Esc` fuerza el retorno al Dashboard del ERP si el flujo automático falla en hacerlo.
* **Manejo de cuelgues o crashes:** Cuando Inmosoft se tilda o se cierra de forma inesperada, el operador debe tomar el control manual:
1. Cancelar el script en la terminal usando `Ctrl + C`.
2. Decidir si cargar manualmente la unidad que quedó a mitad de proceso o editar `liquidacion_template.xlsx` borrando/ocultando las unidades ya procesadas para reiniciar la ejecución desde la propiedad donde cortó.


* **Desfase en la consola:** La terminal imprime el paso siguiente antes de enviarlo a la pantalla. Si Inmosoft se congela, la consola continuará mostrando acciones que no están ocurriendo en el sistema.

---

## 6. Flujo de Trabajo

```mermaid
graph TD
    A[liquidacion_template.xlsx] --> B[Agrupar por Cód. Inmueble con ffill]
    B --> C[Enfocar ventana de Inmosoft]
    
    subgraph SG1 ["Carga de Conceptos (Carga Inmosoft.py)"]
        C --> D[Buscar código de inmueble]
        D --> E[Abrir conceptos de inquilino]
        E --> F[Pegar concepto y monto formateado]
        F --> G[Aplicar cambios]
    end

    subgraph SG2 ["Descarga de Cupones PDF (cupones script.py)"]
        G --> H[Navegar a Útiles -> Descargar PDF]
        H --> I[Pegar ruta del archivo en Save As]
        I --> J[Cerrar emergentes con Esc]
    end
    
    J --> K[Guardado en /CUPONES_DESCARGADOS/]

```

---

## 7. Desafíos Técnicos

### Bug Legacy: Desbordamiento de Búfer (-99999999,00)

El motor gráfico de Inmosoft (32-bit) presenta una falla de análisis numérico cuando se envían caracteres o formatos no sanitizados en la caja de texto `Monto`. Esto provoca un integer overflow asignando el valor límite por defecto `-99999999,00`.

<p align="center">
  <b>1. Intento de carga de monto</b><br>
  <img src="./01-buffer-overflow-input.png" alt="Monto cargado" width="100%" />
</p>

<p align="center">
  <b>2. Error de integer overflow resultante</b><br>
  <img src="./02-buffer-overflow-result.png" alt="Error en el ERP" width="100%" />
</p>

*La sanitización previa del string en el script elimina símbolos monetarios y fuerza la coma decimal antes del pegado, previniendo la corrupción del dato.*
---

## 8. Bugs Conocidos y Cosas a Corregir

1. **Inmosoft se cae solo (fugas de memoria):** Después de 20 o 30 propiedades seguidas, el ERP se cierra de la nada. Hay que volver a abrirlo y reanudar la ejecución desde la fila donde quedó.
2. **Sobreescritura de PDFs:** Si Windows tarda en renderizar el cuadro de "Guardar como", el script envía el `Enter` antes de terminar de pegar el nombre nuevo y lo guarda con el nombre genérico.
3. **Salteo de unidades:** Rara vez se saltea alguna propiedad en el buscador. Pasa cuando el ERP tarda en filtrar la lista y la secuencia tira el clic al menú antes de que termine de cargar.

---

## 9. Mejoras Futuras

1. **Verificación de Estado por Visión Computacional (OpenCV / PyScreeze):** Implementar reconocimiento de patrones visuales para confirmar la presencia de botones ("Agregar", "Aceptar") e íconos de carga antes de enviar clics.
2. **Inspección Dinámica de Elementos (Win32 API / Pywinauto):** Cambiar coordenadas absolutas por la lectura directa de *handles* de ventanas de Windows (`HWND`), detectando cambios de título o controles habilitados.
3. **Manejo de Excepciones por Timeout:** Sustituir los `sleep()` fijos por un patrón *Wait Until Visible* con un tiempo de espera máximo.

---

## Autor

**Teo Quimey Waldemar Londero**

*Analista de Ciberseguridad & Desarrollador RPA*

* [GitHub Profile](https://www.google.com/search?q=https://github.com/tlondero-sec)
* [Portfolio General](https://github.com/tlondero-sec/portfolio)
