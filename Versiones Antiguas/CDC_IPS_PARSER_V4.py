import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "resultado_extraccion_v4.xlsx"

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Índices Base ---
    idx_header = 10         # Fila 11
    idx_base = 12           # Fila 13 (Datos generales)
    
    # Filas donde están los datos mensuales
    idx_mes_indicador = 13  # Fila 14 (Valor Indicador)
    idx_mes_op1 = 15        # Fila 16 (Valor Operando 1)
    idx_mes_op2 = 17        # Fila 18 (Valor Operando 2)

    # --- 3. Extracción Parte 1: Estructura Fija (A-L) ---
    # (Reutilizamos la lógica anterior para A-L)
    encabezados = df_raw.iloc[idx_header, 0:10].tolist()
    valores_base = df_raw.iloc[idx_base, 0:10].tolist()
    
    # Descripciones Operandos (Col K)
    desc_op1 = df_raw.iloc[12, 10]
    desc_op2 = df_raw.iloc[15, 10]
    
    # Estimados Meta (Col L)
    est_meta_op1 = df_raw.iloc[15, 11]
    est_meta_op2 = df_raw.iloc[17, 11]

    # Creamos el DF inicial
    nuevo_df = pd.DataFrame([valores_base], columns=encabezados)
    nuevo_df["Desc. Op1"] = desc_op1
    nuevo_df["Est. Meta Op1"] = est_meta_op1
    nuevo_df["Desc. Op2"] = desc_op2
    nuevo_df["Est. Meta Op2"] = est_meta_op2

    # --- 4. Extracción Parte 2: Los Meses (Ene - Feb) ---
    
    # Diccionario: Nombre del Mes -> Índice de Columna en Excel (A=0)
    # Enero = M (12), Febrero = N (13)
    meses_config = {
        "Ene": 12,
        "Feb": 13
    }

    for mes, col_idx in meses_config.items():
        # Extraemos los 3 valores clave de esa columna
        val_indicador = df_raw.iloc[idx_mes_indicador, col_idx]
        val_op1 = df_raw.iloc[idx_mes_op1, col_idx]
        val_op2 = df_raw.iloc[idx_mes_op2, col_idx]
        
        # Insertamos las 3 nuevas columnas con prefijo del mes
        nuevo_df[f"{mes} Indicador"] = val_indicador
        nuevo_df[f"{mes} Op1"] = val_op1
        nuevo_df[f"{mes} Op2"] = val_op2

    # --- 5. Formato Porcentaje (General) ---
    # Buscamos columnas que parezcan ser porcentajes (Meta, Ponderador, o Indicadores mensuales)
    for col in nuevo_df.columns:
        # Aplicamos formato si es columna de Meta, Ponderador o termina en "Indicador"
        if col in ["Meta 2025", "Ponderador"] or "Indicador" in col:
            nuevo_df[col] = nuevo_df[col].apply(
                lambda x: f"{x:.0%}" if isinstance(x, (int, float)) else x
            )

    # --- 6. Exportar ---
    nuevo_df.to_excel(nombre_archivo_salida, index=False)
    
    print("\n--- Columnas Generadas ---")
    print(nuevo_df.columns.tolist())
    print(f"\n¡Éxito! Archivo guardado como: {nombre_archivo_salida}")

except FileNotFoundError:
    print(f"Error: No se encontró el archivo.")
except Exception as e:
    print(f"Error: {e}")