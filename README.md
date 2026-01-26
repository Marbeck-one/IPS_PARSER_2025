# Sistema de Extracción y Consolidación de Indicadores 2025 (CDC / Riesgos / PMG)

Este proyecto es una herramienta de automatización desarrollada en Python para procesar, limpiar y consolidar planillas de indicadores de gestión (CDC, Riesgos y PMG) que poseen una estructura compleja en Excel (celdas combinadas, encabezados variables).

El script transforma datos visuales no estructurados en bases de datos planas (Tablas) listas para análisis en Power BI, SQL o Excel.

## 🚀 Características Principales

* **Motor Universal:** Detecta automáticamente la estructura de la hoja (CDC, Riesgos o PMG) sin necesidad de configurar filas fijas.
* **Consolidación:** Permite extraer múltiples hojas y guardarlas en un único archivo Excel maestro con pestañas separadas.
* **Limpieza Inteligente:** Estandariza porcentajes (convierte `20%`, `0.2` y `20` a un formato numérico unificado `20.0`).
* **Doble Salida:**
    * **Modo Bruto:** Datos crudos para integración con bases de datos.
    * **Modo Estilizado:** Reportes visuales con formato corporativo (colores, bordes y anchos de columna ajustados).

---

## ⚙️ Cómo Funciona (Flujo Técnico)

El script opera bajo la lógica de **"El Consolidador"**, dividiendo el proceso en 5 etapas secuenciales:

### 1. Interacción y Configuración (`menu_principal`)
El programa inicia actuando como un recepcionista:
1.  **Verificación:** Confirma que el archivo maestro `.xlsx` existe.
2.  **Configuración:** Pregunta al usuario qué formato de salida desea (Bruta, Estilizada o Ambas) y qué hojas desea procesar (CDC, Riesgos, PMG).
3.  **Selección:** Almacena las hojas elegidas en una cola de procesamiento.

### 2. Motor de Extracción (`obtener_dataframe_hoja`)
Se ejecuta una vez por cada hoja seleccionada. Es el cerebro del script:
* **Escaneo Inteligente:** Busca en las primeras 25 filas las palabras clave `NÚMERO` e `INDICADOR` para determinar dónde empieza la tabla, adaptándose si la fila de inicio cambia entre hojas.
* **Mapeo Dinámico:** Identifica en qué columna está cada dato (ej. busca "Ponderador"). Si una columna no existe en una hoja específica (como en Riesgos), el sistema lo nota y rellena con `0` automáticamente.
* **Lógica de Saltos Verticales:** Dado que los Excel originales usan celdas combinadas, el script usa una **Fila Base (`idx`)** y extrae datos relativos:
    * `idx`: Datos generales (Nombre, Fórmula).
    * `idx + 1`: Valores mensuales (% Cumplimiento).
    * `idx + 3`: Operando 1 (Descripción y Valor).
    * `idx + 5`: Operando 2 (Valor).

### 3. Consolidación en Memoria
A diferencia de scripts simples, este no guarda archivos inmediatamente. Almacena cada hoja procesada como un `DataFrame` de Pandas en una lista en la memoria RAM. Esto permite agruparlas más tarde en un solo libro.

### 4. Fabricación del Archivo (`pd.ExcelWriter`)
Una vez todos los datos están listos:
1.  Crea un nuevo archivo Excel (`Planilla_Bruta` o `Planilla_Estilizada`).
2.  Inserta cada `DataFrame` de la memoria en su propia pestaña (Sheet).
3.  Guarda el archivo físico en el disco.

### 5. Maquillaje Visual (`aplicar_estilos_global`)
Si se solicitó la versión estilizada, el script reabre el Excel generado y aplica formato hoja por hoja:
* **Encabezados:** Azul Institucional (`#1F4E78`) con texto blanco.
* **Estructura:** Bordes finos en toda la tabla.
* **Anchos Personalizados:**
    * Columna B (Procesos): Ancho 40.
    * Columnas E-H (Responsables): Ancho 30.
    * Columnas de Meses: Ancho 10.

---

## 📋 Diagrama de Flujo

```mermaid
graph TD
    A[Inicio: Menú Usuario] --> B{¿Archivo Maestro Existe?}
    B -- No --> C[Fin con Error]
    B -- Si --> D[Seleccionar Hojas y Formatos]
    D --> E[Bucle: Procesar cada Hoja]
    E --> F[Detectar Encabezados y Columnas]
    F --> G[Extraer Datos con Saltos Verticales]
    G --> H[Limpiar Porcentajes]
    H --> I[Guardar DataFrame en Memoria]
    I --> E
    E -- Fin Bucle --> J{¿Generar Excel?}
    J --> K[Crear Excel con Pestañas Consolidadas]
    K --> L{¿Es Estilizada?}
    L -- Si --> M[Aplicar Colores, Bordes y Anchos]
    L -- No --> N[Fin]
    M --> N[Fin: Archivos Generados]
