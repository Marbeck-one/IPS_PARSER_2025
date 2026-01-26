import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "resultado_extraccion_v2.xlsx"

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Definición de Índices (Mapeo Excel -> Python) ---
    # Fila 11 Excel (Encabezados) -> índice 10
    # Fila 13 Excel (Datos base y Op1) -> índice 12
    # Fila 16 Excel (Op2) -> índice 15
    idx_encabezados = 10
    idx_fila_base = 12
    idx_fila_op2 = 15
    
    col_k_idx = 10 # La columna K es el índice 10 (A=0, ..., K=10)

    # --- 3. Extracción de Datos ---
    
    # A) Columnas A-J (Base)
    encabezados = df_raw.iloc[idx_encabezados, 0:10].tolist()
    valores_base = df_raw.iloc[idx_fila_base, 0:10].tolist()

    # B) Columnas K (Operandos)
    # Extraemos valores puntuales de la columna K en las filas indicadas
    desc_operando_1 = df_raw.iloc[idx_fila_base, col_k_idx]
    desc_operando_2 = df_raw.iloc[idx_fila_op2, col_k_idx]

    # --- 4. Construcción de la Tabla ---
    
    # Creamos el DataFrame base
    nuevo_df = pd.DataFrame([valores_base], columns=encabezados)

    # Agregamos las nuevas columnas de Operandos
    nuevo_df["Descripcion Operando 1"] = desc_operando_1
    nuevo_df["Descripcion Operando 2"] = desc_operando_2

    # --- 5. Corrección de Formatos (Porcentajes) ---
    indices_porcentaje = ["Meta 2025", "Ponderador"] # Usamos los nombres directamente si ya existen
    
    for col in indices_porcentaje:
        if col in nuevo_df.columns:
            nuevo_df[col] = nuevo_df[col].apply(
                lambda x: f"{x:.0%}" if isinstance(x, (int, float)) else x
            )

    # --- 6. Exportar ---
    nuevo_df.to_excel(nombre_archivo_salida, index=False)
    
    print("\n--- Datos Extraídos (Incluyendo Operandos) ---")
    # Mostramos las columnas nuevas para verificar
    cols_a_mostrar = ["Meta 2025", "Descripcion Operando 1", "Descripcion Operando 2"]
    print(nuevo_df[cols_a_mostrar].T) # .T transpone para leer mejor verticalmente
    
    print(f"\n¡Éxito! Archivo guardado como: {nombre_archivo_salida}")

except FileNotFoundError:
    print(f"Error: No se encontró el archivo '{nombre_archivo_entrada}'.")
except Exception as e:
    print(f"Ocurrió un error: {e}")