---

# Sistema Integral de Gestión de Indicadores 2026 (CDC / Riesgos / PMG) - Motor ETL v7.4.0

Este proyecto es una herramienta de automatización ETL (Extract, Transform, Load) desarrollada en Python, diseñada para procesar, limpiar, estandarizar y transformar planillas de indicadores de gestión complejas provenientes de múltiples fuentes descentralizadas (Regiones, Divisiones, Departamentos).

El sistema actúa como un **puente de datos crítico**, convirtiendo formatos visuales de Excel heterogéneos (con celdas combinadas, encabezados variables, datos dispersos y "escondidos") en estructuras de base de datos relacionales estandarizadas, listas para la carga masiva en el sistema de gestión institucional IPS 2026.

## 🚀 Características Principales

* **Motor de Extracción "Francotirador" (Surgical Extraction):** A diferencia de un lector de Excel tradicional, este motor utiliza una lógica posicional relativa inteligente. Detecta automáticamente el "ancla" de datos (`NÚMERO` e `INDICADOR`) ignorando encabezados institucionales variables, y extrae datos críticos (Metas, Operandos) basándose en su posición relativa (+1 fila, +3 filas, etc.) dentro de bloques visuales complejos.
* **Identificación y Mapeo Inteligente de Responsables:** Infiere automáticamente el Centro de Responsabilidad (CR) propietario y su código interno IP basándose exclusivamente en el nombre del archivo, aplicando reglas de normalización y jerarquía estricta (ej: "Los Rios" -> `DIRECCION REGIONAL DE LOS RIOS`).
* **Consolidación Masiva:** Capaz de procesar más de 27 archivos simultáneamente, unificando datos de CDC, Riesgos y PMG en archivos maestros únicos.
* **Limpieza y Normalización Avanzada:** Estandariza formatos numéricos (miles con punto, decimales con coma/punto), porcentajes, y limpia textos de descripciones (elimina prefijos `(` o sufijos `)*100` residuales).
* **Generación de Paquetes de Carga (Fases 2-5):** Automatiza la creación de las 4 hojas maestras requeridas para la importación del sistema: Variables (`F2`), Variables Aplicadas (`F3`), Indicadores (`F4`) e Indicadores Aplicados (`F5`).
* **Trazabilidad Visual:** Inserta separadores visuales (`--- ORIGEN: Archivo.xlsx ---`) en los archivos de salida para facilitar la auditoría y validación humana de los datos procesados.

---

# 🏗️ FASE 1: Extracción y Estandarización ("El Aplanado")

**Salida:** `1_PLANILLA_SIG_CONSOLIDADO_2026.xlsx` (Hojas: `DATOS_BRUTOS`, `DATOS_ESTILIZADOS`)

En esta etapa, el objetivo es "aplanar" la estructura tridimensional de los Excel originales. El archivo fuente tiene indicadores agrupados en bloques visuales de 6-8 filas. El motor lee estos bloques y los convierte en una única fila horizontal estandarizada por cada indicador.

### 1. Identificación y Metadatos (Fila Base `i`)

El programa escanea el archivo buscando la fila donde aparecen las palabras clave `NÚMERO` e `INDICADOR` (Ancla). Una vez encontrada, itera buscando códigos de indicador (ej. `3.5.1.24`) en la columna A.

| Columna Generada | Fuente en Excel Original | Lógica de Extracción |
| --- | --- | --- |
| **NÚMERO** | Columna A (Fila Base `i`) | ID del indicador. Llave primaria del proceso. |
| **INDICADOR** | Columna B (Fila Base `i`) | Nombre descriptivo del indicador. Se limpia de saltos de línea. |
| **ORIGEN_ARCHIVO** | Nombre del Archivo | Se inyecta para trazabilidad. |
| **RESPONSABLE...** | Nombre del Archivo (Procesado) | Nombre "limpio" del archivo (ej: "CDC Beneficios"). |
| **CODIGO_RESP...** | Inferencia (Mapa Interno) | Código IP asignado según el nombre del archivo (ej: `IP25_712`). |

### 2. Extracción Quirúrgica de Metas y Operandos

Los datos críticos no están en columnas estándar, sino "escondidos" en filas relativas dentro del bloque del indicador.

| Columna Generada | Fuente (Posición Relativa) | Lógica de Extracción ("Francotirador") |
| --- | --- | --- |
| **Meta 2025 (%)** | Fila `i + 1`, Columna E | Busca el valor en la fila siguiente a la base. Convierte porcentajes a decimales/enteros. |
| **Desc. Op1** | Fila `i`, Columna D | Descripción del Numerador. Toma el texto de la fila base. |
| **Desc. Op2** | Fila `i + 3`, Columna D | **Salto:** Baja 3 filas para encontrar la descripción del Denominador. |
| **Est. Meta Op1** | Fila `i + 3`, Columna E | **Salto:** Baja 3 filas. Valor anual estimado para el Numerador. |
| **Est. Meta Op2** | Fila `i + 5`, Columna E | **Salto:** Baja 5 filas. Valor anual estimado para el Denominador. |

### 3. Ciclo Mensual (Octubre - Diciembre)

Dado que las planillas de origen (versión simplificada) suelen traer solo el último trimestre, el sistema extrae los datos reales disponibles y rellena los faltantes.

| Columna Generada | Fuente (Posición Relativa) | Lógica de Extracción |
| --- | --- | --- |
| **Oct Ind (%)** | Fila `i + 1`, Columna F | Valor real del indicador en Octubre. |
| **Oct Op1** | Fila `i + 3`, Columna F | Valor real del Numerador en Octubre. |
| **Oct Op2** | Fila `i + 5`, Columna F | Valor real del Denominador en Octubre. |
| *(Nov y Dic)* | *(Columnas H y J)* | Misma lógica posicional (+1, +3, +5) para Noviembre y Diciembre. |
| **Ene - Sep** | *Inexistente* | Se inyecta valor por defecto `"No aplica"` o `0`. |

### 4. Detección Dinámica de Columnas Opcionales

El script se adapta si el archivo trae o no ciertas columnas.

| Columna Generada | Lógica de Detección |
| --- | --- |
| **UNIDAD_EXTRAIDA** | Busca columna con título "UNIDAD". Si no existe, asigna `"Número"`. |
| **MEDIOS_EXTRAIDOS** | Busca columna "MEDIOS DE VERIFICACIÓN". Si no existe, asigna `"No aplica"`. |

---

# 🏭 FASE 2: Transformación a Variables (`F2`)

**Salida:** `2_CARGA_BRUTA_CONSOLIDADO_2026.xlsx` (Hoja: `F2_VARIABLES`)

Esta fase toma la fila consolidada de la Fase 1 y la **desglosa en dos registros independientes** (`_A` y `_B`) para definir las variables del sistema.

### 1. Separadores de Origen

Inserta una fila visual `--- ORIGEN: Nombre_Archivo.xlsx ---` cada vez que cambia la fuente de datos para mantener el orden.

### 2. Generación de Identificadores y Atributos

| Campo (`cod_interno`) | Lógica de Generación |
| --- | --- |
| **Variable A (Num)** | Código Base + `_A` (ej: `3.5.1.24_A`). |
| **Variable B (Den)** | Código Base + `_B` (ej: `3.5.1.24_B`). |

### 3. Reglas de Negocio Específicas (Mapeo de Columnas)

Se aplican reglas estrictas definidas por el usuario para la configuración de cada variable.

| Columna Excel | Campo Sistema | Valor / Lógica Aplicada |
| --- | --- | --- |
| **E** | `APLICA_DIST_GENERO` | **`?`** (Pendiente de definición manual). |
| **F** | `APLICA_DESP_TERRITORIAL` | **`?`** (Pendiente de definición manual). |
| **G** | `APLICA_SIN_INFORMACION` | **`1`** (Habilitado). |
| **K** | `unidad` | Valor extraído dinámicamente (`UNIDAD_EXTRAIDA`) o "Número". |
| **D** | `medio_verificacion` | Texto extraído (`MEDIOS_EXTRAIDOS`) o "No aplica". |
| **L** | `valor_obligatorio` | **`1`** (Obligatorio). |
| **M** | `permite_medio_escrito` | **`1`** para Variable A / **`0`** para Variable B. |
| **N** | `usa_ultimo_valor_ano` | **`1`** (Habilitado). |

---

# ⚙️ FASE 3: Variables Aplicadas (`F3`)

**Salida:** `2_CARGA_BRUTA_CONSOLIDADO_2026.xlsx` (Hoja: `F3_VAR_APLICADAS`)

Esta etapa asigna las variables creadas en la Fase 2 a los Centros de Responsabilidad correspondientes y configura su comportamiento anual.

### 1. Transformación de Códigos (`cod_var_auto`)

Invierte el sufijo para cumplir la nomenclatura de fórmula del sistema.

* Entrada: `3.5.1.24_A` -> Salida: **`A_3.5.1.24`**

### 2. Asignación de Responsables (Nomenclatura Oficial)

Utiliza el diccionario maestro `MAPA_NOMBRES_OFICIALES` para normalizar el nombre del Centro de Responsabilidad en la columna Q.

| Archivo Origen | Lógica de Normalización (Columna Q) | Resultado |
| --- | --- | --- |
| `CDC REG Los Rios` | Detecta Región + Regla gramatical "DE" | `DIRECCION REGIONAL DE LOS RIOS` |
| `CDC REG Maule` | Detecta Región estándar | `DIRECCION REGIONAL MAULE` |
| `Depto Auditoria` | Detecta Departamento | `DEPARTAMENTO AUDITORIA INTERNA` |
| `SubDir Clientes` | Detecta Subdirección | `SUBDIRECCION SERVICIOS AL CLIENTE` |

### 3. Configuración Técnica

| Columna | Campo | Valor Asignado |
| --- | --- | --- |
| **R** | `cod_region` | **`?`** (Pendiente). |
| **S** | `EMAIL_RESPONSABLE...` | **`prueba@arbol-logika.com`** (Valor por defecto). |
| **T-U** | `EMAILS_REVISORES` | **`NULL`** (Vacíos). |
| **V** | `PERMITE_ADJUNTAR_MEDIO` | **`1`** |
| **W** | `MOSTRAR_TABLA_ANOS` | **`1`** |
| **X** | `FORMULA_VAR_AUTO` | **`SUMA_ANUAL`** |

---

# 📊 FASES 4 y 5: Indicadores Maestros y Aplicación (`F4`, `F5`)

**Salida:** `2_CARGA_BRUTA_CONSOLIDADO_2026.xlsx` (Hojas: `F4_INDICADORES`, `F5_IND_APLICADOS`)

Genera el catálogo de indicadores y su cruce final con metas y responsables.

### Fase 4 (Catálogo)

* Define el indicador con atributos base: `ACTIVO=1`, `UNIDAD=%`, `RANGO_MIN=0`, `RANGO_MAX=100`, `TIPO_META=TOLERANCIA`.
* Extrae Nombre y Descripción limpios de la Fase 1.

### Fase 5 (Aplicación)

* **Cruce Maestro:** Asocia el `INDICADOR_COD` con el `COD_PONDERADO` (Código IP, ej: `IP25_712`) obtenido del mapeo del nombre de archivo.
* **Meta Anual:** Inyecta la meta extraída quirúrgicamente en Fase 1 (`Meta 2025 (%)`).
* **Componentes:** Enlaza las variables A y B generadas (`COMP_A`, `COMP_B`).
* **Fórmulas:** Configura `FORMULA_VAR_AUTO` como `SUMA_ANUAL`.

---

### Resumen del Flujo de Datos Global

1. **Lectura:** El script escanea la carpeta y detecta 27+ archivos.
2. **Identificación:** Por cada archivo, identifica quién es el dueño (Región/División) y cómo se debe llamar oficialmente.
3. **Extracción (F1):** Entra a cada archivo, busca las coordenadas de los datos y extrae la información "sucia".
4. **Transformación (F2-F5):**
* Limpia textos y números.
* Divide indicadores en variables.
* Aplica reglas de negocio (1/0, correos, nombres oficiales).


5. **Carga (Output):** Escribe los 3 archivos Excel finales con formato profesional y separadores de origen.