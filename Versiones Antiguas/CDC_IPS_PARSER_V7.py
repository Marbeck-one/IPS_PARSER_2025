import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "resultado_extraccion_v7_final.xlsx"

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Índices de Filas Clave ---
    idx_header = 10         # Fila 11
    idx_base = 12           # Fila 13 (Datos generales)
    idx_dato_indicador = 13 # Fila 14
    idx_dato_op1 = 15       # Fila 16
    idx_dato_op2 = 17       # Fila 18

    # --- 3. Extracción Base (A - J) ---
    encabezados = df_raw.iloc[idx_header, 0:10].tolist()
    
    # Renombrar Encabezados Clave
    encabezados = [
        "Meta 2025 (%)" if x == "Meta 2025" else 
        "Ponderador (%)" if x == "Ponderador" else x 
        for x in encabezados
    ]

    valores_base = df_raw.iloc[idx_base, 0:10].tolist()
    
    # --- 4. Extracción Operandos (K - L) ---
    desc_op1 = df_raw.iloc[idx_base, 10]
    desc_op2 = df_raw.iloc[15, 10]
    est_meta_op1 = df_raw.iloc[15, 11]
    est_meta_op2 = df_raw.iloc[17, 11]

    # Crear DF Inicial
    nuevo_df = pd.DataFrame([valores_base], columns=encabezados)
    nuevo_df["Desc. Op1"] = desc_op1
    nuevo_df["Est. Meta Op1"] = est_meta_op1
    nuevo_df["Desc. Op2"] = desc_op2
    nuevo_df["Est. Meta Op2"] = est_meta_op2

    # --- 5. Extracción Meses y Proyecciones ---
    mapa_columnas = {
        "Ene": 12, "Feb": 13, "Mar": 15, "Abr": 17, "May": 19, "Jun": 21,
        "Jul": 23, "Ago": 25, "Sept": 27, "Oct": 29, "Nov": 31, "Dic": 33,
        "Cump. Proy.": 34
    }

    for nombre, col_idx in mapa_columnas.items():
        val_ind = df_raw.iloc[idx_dato_indicador, col_idx]
        val_op1 = df_raw.iloc[idx_dato_op1, col_idx]
        val_op2 = df_raw.iloc[idx_dato_op2, col_idx]
        
        # Agregamos columnas con el sufijo (%) para el indicador
        nuevo_df[f"{nombre} Ind (%)"] = val_ind
        nuevo_df[f"{nombre} Op1"] = val_op1
        nuevo_df[f"{nombre} Op2"] = val_op2

    # --- 6. Columnas Finales ---
    # AJ: Cumplimiento Meta
    nuevo_df["Cumplimiento Meta (%)"] = df_raw.iloc[idx_dato_op1, 35]
    
    # Textos finales
    nuevo_df["Medios Verificación"] = df_raw.iloc[idx_base, 36]
    nuevo_df["Control Cambios"] = df_raw.iloc[idx_base, 37]
    nuevo_df["Instrumentos Gestión"] = df_raw.iloc[idx_base, 38]

    # --- 7. Limpieza y Formato Numérico ---
    
    # Función para limpiar cualquier residuo (ej: si Excel trae texto "20%")
    def limpiar_numero(val):
        if pd.isna(val) or val == "":
            return 0
        if isinstance(val, (int, float)):
            return val
        if isinstance(val, str):
            try:
                # Quitamos % y cambiamos coma por punto
                val_limpio = val.replace('%', '').replace(',', '.')
                return float(val_limpio) / 100 # Dividimos por 100 para volver a decimal (20 -> 0.2)
            except:
                return 0
        return 0

    # Identificar columnas de porcentaje para asegurar limpieza
    cols_porcentaje = [c for c in nuevo_df.columns if "(%)" in c]
    
    for col in cols_porcentaje:
        nuevo_df[col] = nuevo_df[col].apply(limpiar_numero)

    # Rellenar cualquier otro NaN con 0
    nuevo_df = nuevo_df.fillna(0)

    # --- 8. Exportar ---
    nuevo_df.to_excel(nombre_archivo_salida, index=False)
    print(f"¡Éxito! Archivo guardado como: {nombre_archivo_salida}")

except Exception as e:
    print(f"Ocurrió un error: {e}")