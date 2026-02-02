
---

# Sistema Integral de Gestión de Indicadores 2026 (CDC / Riesgos / PMG)

Este proyecto es una herramienta de automatización desarrollada en Python diseñada para procesar, limpiar y transformar planillas de indicadores de gestión complejas.

El sistema actúa como un **puente de datos**, convirtiendo formatos visuales de Excel (celdas combinadas, encabezados variables) en estructuras de base de datos estandarizadas para el sistema IPS 2026.

## 🚀 Características Principales

* **Motor de Extracción Universal:** Detecta automáticamente la estructura de hojas (CDC, Riesgos, PMG) sin configuración manual de filas.
* **Consolidación Inteligente:** Agrupa múltiples fuentes en archivos maestros.
* **Limpieza de Datos:** Estandariza porcentajes y limpia textos (elimina fórmulas o paréntesis residuales en descripciones).
* **Módulo de Variables (Fase 2):** Desglosa cada indicador en sus componentes variables (Numerador/Denominador) generando filas `_A` y `_B`.
* **Módulo de Variables Aplicadas (Fase 3):** Genera la configuración anual, transformando sufijos en prefijos y asignando parámetros de control y correos.
* **Doble Salida (Dual):** Genera archivos con pestañas separadas para:
* **Bruta:** Datos puros para integraciones.
* **Estilizada:** Formato visual para revisión humana.



---

# 🏗️ FASE 1: Extracción y Estandarización

**Salida:** `Planilla_Bruta_2025.xlsx` y `Planilla_Estilizada_2025.xlsx`

En esta etapa, el objetivo es "aplanar" el Excel original. El archivo original tiene una estructura tridimensional compleja (celdas combinadas que agrupan filas). El programa lee bloques verticales y los convierte en una sola fila horizontal por indicador.

### 1. Identificación y Metadatos (Datos de la Fila Base)

El programa escanea la columna "NÚMERO". Cuando encuentra un código (ej. `5.4.1.61`), marca esa fila como **Fila Base (`idx`)**.

| Columna Generada | Fuente en Excel Original | Lógica de Extracción |
| --- | --- | --- |
| **NÚMERO** | Columna "NÚMERO" (Fila Base) | Es el ID del indicador. Se usa como ancla para todo el proceso. |
| **PRODUCTO O PROCESO...** | Columna "PRODUCTO..." (Fila Base) | Extrae el texto descriptivo del proceso macro. |
| **INDICADOR** | Columna "INDICADOR" (Fila Base) | El nombre principal del indicador. |
| **FORMULA** | Columna "FORMULA" (Fila Base) | La fórmula matemática textual. |
| **UNIDAD** | Columna "UNIDAD" (Fila Base) | La unidad de medida (ej. "Porcentaje", "Número"). |
| **RESPONSABLE...** | Columna "RESPONSABLE..." (Fila Base) | Nombre de la jefatura o área responsable. |
| **GESTOR** | Columna "GESTOR" (Fila Base) | Persona operativa a cargo. |
| **SUPERVISORES** | Columna "SUPERVISORES" (Fila Base) | Quien supervisa la gestión. |

### 2. Metas y Ponderadores (Datos Estratégicos)

Estos datos suelen estar en la misma fila base o cerca de ella.

| Columna Generada | Fuente en Excel Original | Lógica de Extracción |
| --- | --- | --- |
| **Meta 2025 (%)** | Columna "Meta 2025" (Fila Base) | Se limpia: si es `1` se convierte a `100`, si es `0.9` a `90`. |
| **Ponderador (%)** | Columna "Ponderador" (Fila Base) | **Lógica Especial:** Si la hoja es "Riesgos" o "PMG" (donde esta columna no existe), el programa inserta automáticamente un **0**. En CDC extrae el valor real. |

### 3. Definición de Operandos (El "Diccionario" de la Fórmula)

Aquí el programa debe "saltar" filas hacia abajo desde la Fila Base (`idx`) para encontrar las definiciones.

| Columna Generada | Fuente en Excel Original | Lógica de Extracción (Saltos) |
| --- | --- | --- |
| **Desc. Op1** | Columna "Operandos" (Fila Base) | Toma el texto de la misma fila del indicador. Describe el Numerador. |
| **Desc. Op2** | Columna "Operandos" (**Fila Base + 3**) | **Salto:** Baja 3 filas para encontrar la descripción del Denominador. |
| **Est. Meta Op1** | Columna "Operandos Est." (**Fila Base + 3**) | **Salto:** Baja 3 filas. Es el valor numérico estimado para el Numerador. |
| **Est. Meta Op2** | Columna "Operandos Est." (**Fila Base + 5**) | **Salto:** Baja 5 filas. Es el valor numérico estimado para el Denominador. |

### 4. Ciclo Mensual (Enero a Diciembre)

El programa itera por cada mes (columnas Ene, Feb, Mar...). Para *cada mes*, extrae un trío de datos vertical.

*Ejemplo para Enero:*

| Columna Generada | Fuente en Excel Original | Lógica de Extracción (Coordenadas) |
| --- | --- | --- |
| **Ene Ind (%)** | Columna "Ene." (**Fila Base + 1**) | Es el % de cumplimiento del mes. Se limpia matemáticamente. |
| **Ene Op1** | Columna "Ene." (**Fila Base + 3**) | Es el valor real ejecutado del Numerador en Enero. |
| **Ene Op2** | Columna "Ene." (**Fila Base + 5**) | Es el valor real ejecutado del Denominador en Enero. |

*(Esta lógica se repite idéntica para Feb, Mar, Abr... hasta Dic).*

### 5. Proyecciones y Cierres

Datos ubicados al final de la tabla horizontal.

| Columna Generada | Fuente en Excel Original | Lógica de Extracción |
| --- | --- | --- |
| **Cump. Proy. Ind (%)** | Columna "Cumplimiento Proy." (**Fila + 1**) | Proyección del indicador a fin de año. |
| **Cump. Proy. Op1** | Columna "Cumplimiento Proy." (**Fila + 3**) | Proyección del Numerador. |
| **Cump. Proy. Op2** | Columna "Cumplimiento Proy." (**Fila + 5**) | Proyección del Denominador. |
| **Cumplimiento Meta (%)** | Columna "% Cump. Meta" (**Fila + 3**) | Porcentaje final de logro respecto a la meta. |
| **Medios Verificación** | Columna "Medios..." (Fila Base) | Texto largo con la evidencia requerida. |
| **Control Cambios** | Columna "Control..." (Fila Base) | Historial de modificaciones. |
| **Instrumentos Gestión** | Columna "Instrumentos..." (Fila Base) | Documentos asociados. |

---

# 🏭 FASE 2: Transformación a Variables IPS

**Salida:** `VARIABLES_IPS_2026.xlsx`

Esta fase toma la fila "aplanada" de la Fase 1 y la **divide en dos filas independientes** (`_A` y `_B`) para alimentar el sistema de carga masiva.

### 1. Separadores de Sección

Antes de procesar los datos, el sistema inserta una "Fila Título" para separar CDC, Riesgos y PMG visualmente.

* **Columna A:** `--- CDC VARIABLES ---`
* **Resto:** Vacío.

### 2. Generación de Identificadores (`cod_interno`)

El sistema analiza la columna `NÚMERO` de la Fase 1.

| Columna A (cod_interno) | Lógica del Programa |
| --- | --- |
| **Fila A (Numerador)** | Toma el código original y agrega `_A`. <br>

<br>

<br>Ej: `5.4.1.61` ➔ **`5.4.1.61_A`** |
| **Fila B (Denominador)** | Toma el código original y agrega `_B`. <br>

<br>

<br>Ej: `5.4.1.61` ➔ **`5.4.1.61_B`** |
| *Caso Especial: Nuevos* | Si el código original está vacío o dice "INDICADOR NUEVO", genera un ID secuencial único para evitar errores.<br>

<br>

<br>Ej: `INDICADOR_NUEVO_1_A_CDC`. |

### 3. Limpieza de Textos (`nombre_variable` y `descripcion`)

El sistema limpia "basura" sintáctica que viene del Excel original.

| Columna B y C | Fuente (Fase 1) | Algoritmo de Limpieza |
| --- | --- | --- |
| **Fila A** | `Desc. Op1` | **Regex:** Busca si el texto empieza con `(`. Si es así, lo elimina.<br>

<br>

<br>Original: `(Sumatoria de hitos...`<br>

<br>

<br>Final: `Sumatoria de hitos...` |
| **Fila B** | `Desc. Op2` | **Regex:** Busca si el texto termina con `)*100`. Si es así, lo elimina.<br>

<br>

<br>Original: `...total de hitos)*100`<br>

<br>

<br>Final: `...total de hitos` |

### 4. Asignación de Verificadores

| Columna D | Fuente (Fase 1) | Lógica |
| --- | --- | --- |
| **medio_verificacion** | `Medios Verificación` | Se copia el **mismo texto** tanto para la fila A como para la fila B. Ambas variables comparten el mismo medio de prueba. |

### 5. Banderas de Configuración (Hardcoded)

Estas columnas tienen valores fijos definidos por tus reglas de negocio ("Hardcoded" significa que el código siempre pone el mismo valor, no lo lee del Excel).

| Columna | Título | Valor Asignado | Significado Técnico |
| --- | --- | --- | --- |
| **E** | `APLICA_DIST_GENERO` | **0** | No requiere distinción hombre/mujer. |
| **F** | `APLICA_DESP_TERRITORIAL` | **0** | No requiere desglose regional. |
| **G** | `APLICA_SIN_INFORMACION` | **1** | Permite reportar "Sin Información". |
| **H** | `APLICA_VAL_PERS_JUR` | **0** | No aplica a personas jurídicas. |
| **I** | `requiere_medio` | **0** | (Regla específica del negocio). |
| **J** | `texto_ayuda` | **NULL** (Vacío) | Campo opcional dejado en blanco. |
| **K** | `unidad` | **NULL** (Vacío) | Campo opcional dejado en blanco. |
| **L** | `valor_obligatorio` | **1** | El sistema exigirá que este dato no esté vacío. |
| **M** | `permite_medio_escrito` | **1** | Permite ingresar observaciones de texto. |
| **N** | `usa_ultimo_valor_ano` | **1** | Configuración de arrastre de datos anuales. |

---

# ⚙️ FASE 3: Generación de Variables Aplicadas

**Salida:** `VARIABLES_APLICADAS_IPS_2026.xlsx`

Esta etapa final genera la planilla de configuración anual para el sistema, utilizando como base los datos consolidados de la Fase 2.

### 1. Transformación de Códigos (`cod_var_auto`)

El sistema toma los códigos generados en la Fase 2 y aplica una transformación de **Sufijo a Prefijo** para cumplir con la nomenclatura interna del sistema IPS.

| Código Fase 2 (Entrada) | Transformación | Código Fase 3 (Salida) |
| --- | --- | --- |
| `5.4.1.61_A` | Sufijo `_A` pasa al inicio | **`A_5.4.1.61`** |
| `5.4.1.61_B` | Sufijo `_B` pasa al inicio | **`B_5.4.1.61`** |
| `INDICADOR_NUEVO_1_A_CDC` | Se reordena la letra | **`A_INDICADOR_NUEVO_1_CDC`** |

### 2. Configuración de Vigencia y Meses

Se establecen los parámetros temporales de la variable.

| Columna | Nombre Campo | Valor Asignado | Descripción |
| --- | --- | --- | --- |
| **C** | `ano_mes_ini` | **202501** | Inicio de vigencia: Enero 2025. |
| **D** | `ano_mes_fin` | **202512** | Fin de vigencia: Diciembre 2025. |
| **E - P** | `ENE` ... `DIC` | **1** | Bandera (1) que activa la variable para cada mes del año. |

### 3. Asignación de Responsables y Correos

Se configuran los correos electrónicos para el flujo de aprobación y carga.

| Columna | Nombre Campo | Valor Asignado | Nota |
| --- | --- | --- | --- |
| **S** | `EMAIL_RESPONSABLE` | `prueba@arbol-logika.com` | Correo por defecto para pruebas de carga. |
| **T** | `EMAIL_PRIMER_REV` | **NULL** (Vacío) | Se deja en blanco intencionalmente. |
| **U** | `EMAIL_SEGUNDO_REV` | **NULL** (Vacío) | Se deja en blanco intencionalmente. |

### 4. Parámetros Técnicos Adicionales

| Columna | Nombre Campo | Valor Asignado | Descripción |
| --- | --- | --- | --- |
| **Q** | `cod_centro_resp...` | **NULL** (Vacío) | Centro de responsabilidad (pendiente de asignar). |
| **R** | `cod_region` | **NULL** (Vacío) | Código regional (pendiente de asignar). |
| **V** | `PERMITE_ADJUNTAR` | **1** | Habilita la subida de archivos adjuntos. |
| **W** | `MOSTRAR_TABLA` | **1** | Visualización de tabla histórica. |
| **X** | `FORMULA_VAR_AUTO` | **SUMA_ANUAL** | Fórmula de cálculo automático. |

### 5. Preservación de Estructura Visual

El sistema respeta los separadores de sección (`--- CDC VARIABLES ---`) generados en la Fase 2 y les aplica formato de **Negrita** en el Excel final para mantener la legibilidad por grupos (CDC, Riesgos, PMG).

---

### Resumen del Flujo de Datos Global

1. **Excel Original:** Datos en "bloques" 3D.
⬇️ *Parser Fase 1*
2. **Planilla Bruta:** Una fila larga por indicador.
⬇️ *Transformador Fase 2*
3. **Variables IPS:** Desglose en filas A/B + Limpieza.
⬇️ *Aplicador Fase 3*
4. **Variables Aplicadas:** Prefijos, Correos y Configuración Anual.
