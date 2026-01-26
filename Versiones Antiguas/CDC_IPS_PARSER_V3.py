import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "resultado_extraccion_v3.xlsx"

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Definición de Índices (Mapa del Tesoro) ---
    # Indices basados en 0 (Fila Excel - 1)
    
    # Encabezados
    idx_header = 10         # Fila 11
    
    # Datos Base
    idx_base = 12           # Fila 13
    
    # Columna K (Descripciones)
    idx_desc_op1 = 12       # Fila 13
    idx_desc_op2 = 15       # Fila 16
    
    # Columna L (Valores Estimados)
    idx_val_op1 = 15        # Fila 16 (Donde está el valor real del Op1)
    idx_val_op2 = 17        # Fila 18 (Donde está el valor real del Op2)
    
    # Indices de Columnas
    col_k = 10
    col_l = 11

    # --- 3. Extracción ---
    
    # A) Base (A-J)
    encabezados = df_raw.iloc[idx_header, 0:10].tolist()
    valores_base = df_raw.iloc[idx_base, 0:10].tolist()

    # B) Columna K (Descripciones)
    desc_op1 = df_raw.iloc[idx_desc_op1, col_k]
    desc_op2 = df_raw.iloc[idx_desc_op2, col_k]

    # C) Columna L (Valores Estimados)
    val_est_op1 = df_raw.iloc[idx_val_op1, col_l]
    val_est_op2 = df_raw.iloc[idx_val_op2, col_l]

    # --- 4. Armado de la Tabla ---
    nuevo_df = pd.DataFrame([valores_base], columns=encabezados)

    # Agregamos lo nuevo
    nuevo_df["Descripcion Operando 1"] = desc_op1
    nuevo_df["Operando 1 estimado meta"] = val_est_op1
    
    nuevo_df["Descripcion Operando 2"] = desc_op2
    nuevo_df["Operando 2 estimado meta"] = val_est_op2

    # --- 5. Corrección de Formato (Porcentajes) ---
    cols_porcentaje = ["Meta 2025", "Ponderador"]
    for col in cols_porcentaje:
        if col in nuevo_df.columns:
            nuevo_df[col] = nuevo_df[col].apply(
                lambda x: f"{x:.0%}" if isinstance(x, (int, float)) else x
            )

    # --- 6. Exportar ---
    nuevo_df.to_excel(nombre_archivo_salida, index=False)
    
    print("\n--- Vista Previa de las Nuevas Columnas ---")
    cols_nuevas = ["Descripcion Operando 1", "Operando 1 estimado meta", 
                   "Descripcion Operando 2", "Operando 2 estimado meta"]
    print(nuevo_df[cols_nuevas].T)
    
    print(f"\n¡Listo! Archivo generado: {nombre_archivo_salida}")

except FileNotFoundError:
    print(f"Error: No encuentra '{nombre_archivo_entrada}'.")
except Exception as e:
    print(f"Error: {e}")