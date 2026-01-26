import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "resultado_extraccion_v6_final.xlsx"

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Índices de Filas Clave ---
    idx_header = 10         # Fila 11
    idx_base = 12           # Fila 13 (Texto base, Descripciones, Medios Verif)
    
    # Filas para datos variables (Meses y Cumplimiento Proyectado)
    idx_dato_indicador = 13  # Fila 14
    idx_dato_op1 = 15        # Fila 16
    idx_dato_op2 = 17        # Fila 18

    # --- 3. Extracción Parte 1: Estructura Fija Inicial (A - L) ---
    
    # Columnas A-J (Base)
    encabezados = df_raw.iloc[idx_header, 0:10].tolist()
    valores_base = df_raw.iloc[idx_base, 0:10].tolist()
    
    # Columnas K-L (Operandos Definición)
    desc_op1 = df_raw.iloc[idx_base, 10]      # K13
    desc_op2 = df_raw.iloc[15, 10]            # K16
    est_meta_op1 = df_raw.iloc[15, 11]        # L16 (Valor real)
    est_meta_op2 = df_raw.iloc[17, 11]        # L18 (Valor real)

    # Creamos DF Inicial
    nuevo_df = pd.DataFrame([valores_base], columns=encabezados)
    nuevo_df["Desc. Op1"] = desc_op1
    nuevo_df["Est. Meta Op1"] = est_meta_op1
    nuevo_df["Desc. Op2"] = desc_op2
    nuevo_df["Est. Meta Op2"] = est_meta_op2

    # --- 4. Extracción Parte 2: Meses + Cumplimiento Proyectado (AI) ---
    # Tratamos a "Cumplimiento Proyectado" (AI) como un mes más porque tiene la misma estructura
    
    mapa_columnas_complejas = {
        "Ene": 12, "Feb": 13, "Mar": 15, "Abr": 17, "May": 19, "Jun": 21,
        "Jul": 23, "Ago": 25, "Sept": 27, "Oct": 29, "Nov": 31, "Dic": 33,
        "Cump. Proy.": 34  # Columna AI
    }

    print("Procesando Meses y Cumplimiento Proyectado...")
    for nombre, col_idx in mapa_columnas_complejas.items():
        # Extraemos el trío de valores
        val_ind = df_raw.iloc[idx_dato_indicador, col_idx]
        val_op1 = df_raw.iloc[idx_dato_op1, col_idx]
        val_op2 = df_raw.iloc[idx_dato_op2, col_idx]
        
        # Agregamos al DF
        nuevo_df[f"{nombre} Ind"] = val_ind
        nuevo_df[f"{nombre} Op1"] = val_op1
        nuevo_df[f"{nombre} Op2"] = val_op2

    # --- 5. Extracción Parte 3: Columnas Finales (AJ - AM) ---
    # Estas tienen comportamientos únicos
    
    # AJ: % Cumplimiento de Meta (Dato en fila 16/idx 15)
    nuevo_df["% Cumplimiento Meta"] = df_raw.iloc[idx_dato_op1, 35]
    
    # AK, AL, AM: Textos descriptivos (Dato en fila 13/idx 12)
    nuevo_df["Medios Verificación"] = df_raw.iloc[idx_base, 36]
    nuevo_df["Control Cambios"] = df_raw.iloc[idx_base, 37]
    nuevo_df["Instrumentos Gestión"] = df_raw.iloc[idx_base, 38]

    # --- 6. Formato Porcentaje ---
    print("Aplicando formatos finales...")
    cols_porcentaje = ["Meta 2025", "Ponderador", "% Cumplimiento Meta"]
    
    for col in nuevo_df.columns:
        es_pct = False
        if col in cols_porcentaje: es_pct = True
        if "Ind" in col and "Indicador" not in col: es_pct = True # Para Ene Ind, Feb Ind...
        
        if es_pct:
            nuevo_df[col] = nuevo_df[col].apply(
                lambda x: f"{x:.0%}" if isinstance(x, (int, float)) else x
            )

    # --- 7. Exportar ---
    nuevo_df.to_excel(nombre_archivo_salida, index=False)
    print(f"¡Misión cumplida! Archivo completo guardado como: {nombre_archivo_salida}")

except FileNotFoundError:
    print("Error: Archivo no encontrado.")
except Exception as e:
    print(f"Error inesperado: {e}")