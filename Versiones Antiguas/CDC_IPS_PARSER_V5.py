import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "resultado_extraccion_v5_anual.xlsx"

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Índices de Filas (Vertical) ---
    idx_header = 10         # Fila 11
    idx_base = 12           # Fila 13 (Datos generales y Descripciones)
    
    # Filas de datos mensuales
    idx_mes_indicador = 13  # Fila 14 (Valor Indicador)
    idx_mes_op1 = 15        # Fila 16 (Valor Operando 1)
    idx_mes_op2 = 17        # Fila 18 (Valor Operando 2)

    # --- 3. Extracción Parte 1: Base (Columnas A - L) ---
    
    # A) Base (A-J)
    encabezados = df_raw.iloc[idx_header, 0:10].tolist()
    valores_base = df_raw.iloc[idx_base, 0:10].tolist()
    
    # B) Descripciones y Estimados (K y L)
    # Col K (Indice 10) y Col L (Indice 11)
    desc_op1 = df_raw.iloc[12, 10]
    desc_op2 = df_raw.iloc[15, 10]
    est_meta_op1 = df_raw.iloc[15, 11] # Valor real fila 16
    est_meta_op2 = df_raw.iloc[17, 11] # Valor real fila 18

    # Creamos DF
    nuevo_df = pd.DataFrame([valores_base], columns=encabezados)
    nuevo_df["Desc. Op1"] = desc_op1
    nuevo_df["Est. Meta Op1"] = est_meta_op1
    nuevo_df["Desc. Op2"] = desc_op2
    nuevo_df["Est. Meta Op2"] = est_meta_op2

    # --- 4. Extracción Parte 2: Todo el Año (Sin Acumulados) ---
    
    # Mapa exacto de columnas para cada mes (Saltando Acumulados)
    # Ene=12, Feb=13 (Sin acum entre ellos)
    # Mar=15 (Saltamos 14 Acum Feb)
    # Abr=17 (Saltamos 16 Acum Mar) ... y así sucesivamente (+2)
    mapa_meses = {
        "Ene": 12, "Feb": 13, "Mar": 15, "Abr": 17, "May": 19, "Jun": 21,
        "Jul": 23, "Ago": 25, "Sept": 27, "Oct": 29, "Nov": 31, "Dic": 33
    }

    print("\nProcesando meses...")
    for mes, col_idx in mapa_meses.items():
        # Extraemos el trío
        val_indicador = df_raw.iloc[idx_mes_indicador, col_idx]
        val_op1 = df_raw.iloc[idx_mes_op1, col_idx]
        val_op2 = df_raw.iloc[idx_mes_op2, col_idx]
        
        # Insertamos columnas
        # Nota: Usamos nombres cortos para que el Excel no sea gigante
        nuevo_df[f"{mes} Ind"] = val_indicador
        nuevo_df[f"{mes} Op1"] = val_op1
        nuevo_df[f"{mes} Op2"] = val_op2

    # --- 5. Formatos ---
    print("Aplicando formatos...")
    for col in nuevo_df.columns:
        # Formateamos si es Meta, Ponderador o cualquier columna de "Indicador"
        es_porcentaje = False
        if col in ["Meta 2025", "Ponderador"]: es_porcentaje = True
        if "Ind" in col and "Indicador" not in col: es_porcentaje = True # Para los meses "Ene Ind", etc.
        
        if es_porcentaje:
            nuevo_df[col] = nuevo_df[col].apply(
                lambda x: f"{x:.0%}" if isinstance(x, (int, float)) else x
            )

    # --- 6. Exportar ---
    nuevo_df.to_excel(nombre_archivo_salida, index=False)
    print(f"¡Listo! Archivo anual generado: {nombre_archivo_salida}")

except FileNotFoundError:
    print("Error: No se encuentra el archivo.")
except Exception as e:
    print(f"Error: {e}")