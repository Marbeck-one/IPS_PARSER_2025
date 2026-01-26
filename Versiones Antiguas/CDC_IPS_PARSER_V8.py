import pandas as pd
import os

# --- 1. Configuración ---
nombre_archivo_entrada = "Proyecciones Indicadores 2025 - División Planificación (1).xlsx"
nombre_hoja = "CDC 2025"
nombre_archivo_salida = "Planilla_Indicadores_Procesada_Final.xlsx"

# Función maestra de limpieza de porcentajes (0-100)
def limpiar_porcentaje_real(val):
    """Convierte 1.0 -> 100, 0.2 -> 20, '20%' -> 20"""
    if pd.isna(val) or val == "":
        return 0
    
    # Caso 1: Es texto (ej: "20%" o "100%")
    if isinstance(val, str):
        # Quitamos % y espacios, cambiamos coma por punto
        limpio = val.replace('%', '').replace(',', '.').strip()
        try:
            return float(limpio) # "20" pasa a ser 20.0
        except:
            return 0
            
    # Caso 2: Es número (Excel guarda 100% como 1.0, 20% como 0.2)
    if isinstance(val, (int, float)):
        return val * 100  # Escalamos a 0-100
        
    return 0

try:
    print(f"Leyendo archivo: {nombre_archivo_entrada}...")
    df_raw = pd.read_excel(nombre_archivo_entrada, sheet_name=nombre_hoja, header=None)

    # --- 2. Detección de Filas de Inicio ---
    # Buscamos todas las filas donde la Columna A (índice 0) tenga un valor (ej: "5.4.1.61")
    # Empezamos a buscar desde la fila 11 (índice 11) en adelante
    indices_inicio = []
    for i in range(11, len(df_raw)):
        val = df_raw.iloc[i, 0]
        # Si no es nulo y no es el encabezado "NÚMERO"
        if pd.notna(val) and str(val).strip() != "" and str(val) != "NÚMERO":
            indices_inicio.append(i)
    
    print(f"Se encontraron {len(indices_inicio)} indicadores para procesar.")
    
    # --- 3. Bucle de Procesamiento ---
    lista_filas_procesadas = []
    
    # Encabezados Base (tomados de la fila 10 fija)
    encabezados_raw = df_raw.iloc[10, 0:10].tolist()
    # Ajustamos nombres de encabezados
    encabezados_finales = [
        "Meta 2025 (%)" if x == "Meta 2025" else 
        "Ponderador (%)" if x == "Ponderador" else x 
        for x in encabezados_raw
    ]

    for idx in indices_inicio:
        # Definimos los offsets relativos (basados en el análisis anterior)
        # idx = Fila Base (ej: 12)
        # idx + 1 = Fila Indicador Mensual (ej: 13)
        # idx + 3 = Fila Op1 / Cumplimiento Meta (ej: 15)
        # idx + 5 = Fila Op2 (ej: 17)
        
        # A) Datos Base (A-J)
        fila_base = df_raw.iloc[idx, 0:10].tolist()
        datos_dict = dict(zip(encabezados_finales, fila_base))
        
        # B) Operandos (K-L)
        datos_dict["Desc. Op1"] = df_raw.iloc[idx, 10]      # Col K, Fila Base
        datos_dict["Desc. Op2"] = df_raw.iloc[idx+3, 10]    # Col K, Fila Base+3
        datos_dict["Est. Meta Op1"] = df_raw.iloc[idx+3, 11] # Col L, Fila Base+3
        datos_dict["Est. Meta Op2"] = df_raw.iloc[idx+5, 11] # Col L, Fila Base+5
        
        # C) Meses y Proyección (Columnas M en adelante)
        mapa_cols = {
            "Ene": 12, "Feb": 13, "Mar": 15, "Abr": 17, "May": 19, "Jun": 21,
            "Jul": 23, "Ago": 25, "Sept": 27, "Oct": 29, "Nov": 31, "Dic": 33,
            "Cump. Proy.": 34
        }
        
        for mes, col_idx in mapa_cols.items():
            datos_dict[f"{mes} Ind (%)"] = df_raw.iloc[idx+1, col_idx] # Fila Base+1
            datos_dict[f"{mes} Op1"] = df_raw.iloc[idx+3, col_idx]     # Fila Base+3
            datos_dict[f"{mes} Op2"] = df_raw.iloc[idx+5, col_idx]     # Fila Base+5
            
        # D) Columnas Finales (AJ - AM)
        datos_dict["Cumplimiento Meta (%)"] = df_raw.iloc[idx+3, 35] # AJ, Fila Base+3
        datos_dict["Medios Verificación"] = df_raw.iloc[idx, 36]     # AK, Fila Base
        datos_dict["Control Cambios"] = df_raw.iloc[idx, 37]         # AL, Fila Base
        datos_dict["Instrumentos Gestión"] = df_raw.iloc[idx, 38]    # AM, Fila Base
        
        lista_filas_procesadas.append(datos_dict)

    # --- 4. Creación de Tabla y Limpieza Final ---
    df_final = pd.DataFrame(lista_filas_procesadas)
    
    # Rellenar vacíos con 0 antes de procesar números
    df_final = df_final.fillna(0)
    
    # Aplicar conversión de porcentaje (0-100) a las columnas correspondientes
    cols_porcentaje = [c for c in df_final.columns if "(%)" in c]
    
    print("Aplicando conversión a escala 0-100...")
    for col in cols_porcentaje:
        df_final[col] = df_final[col].apply(limpiar_porcentaje_real)

    # --- 5. Exportar ---
    df_final.to_excel(nombre_archivo_salida, index=False)
    print(f"¡Proceso completado! Archivo maestro guardado como: {nombre_archivo_salida}")
    
    # Vista previa
    cols_check = ["NÚMERO", "Meta 2025 (%)", "Ene Ind (%)", "Ponderador (%)"]
    print("\nPrimeras filas (verificación de escala):")
    print(df_final[cols_check].head())

except Exception as e:
    print(f"Error fatal: {e}")