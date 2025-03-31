# -*- coding: utf-8 -*-
"""
Script de Análisis de Emisiones Chilexpress - Predictivo, Simulación y Cumplimiento SBTi
Autor: Asistente AI
Requerimientos: Python 3.8+, pandas, matplotlib, scikit-learn, openpyxl
"""

# ======================================
# Configuración Inicial y Librerías
# ======================================
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import os
import sys
import re

# Configurar rutas específicas
DIR_PROYECTO = r"C:\Users\Alexander\OneDrive\Escritorio\portafolio_chilexpress\Proyecto1_Limpieza"
ARCHIVO_EXCEL_ENTRADA = os.path.join(DIR_PROYECTO, "emisiones_chilexpress.xlsx")
HOJA_EXCEL = "hoja1"  # Cambiar si es necesario
# <<< NUEVO: Nombre del archivo Excel de salida para los resultados >>>
ARCHIVO_RESULTADOS_EXCEL = os.path.join(DIR_PROYECTO, "resultados_analisis_emisiones.xlsx")

# ======================================
# Funciones de Carga y Validación
# ======================================
def cargar_datos():
    """Carga y valida estructura del Excel."""
    cols_requeridas = [
        'Categoría', 'Subcategoría', 'Emisiones 2021 (tCO₂e)',
        'Emisiones 2022 Real (tCO₂e)', 'Metodología/Factor Emisión',
        'Notas Ajustadas', 'Variación (%)', 'Indicadores por pieza 2022 (kgCO₂e)'
    ]
    try:
        # Leer Excel ignorando filas vacías iniciales y buscando la cabecera correcta
        # Intentaremos encontrar la fila que contiene 'Categoría' como inicio de la cabecera
        xls = pd.ExcelFile(ARCHIVO_EXCEL_ENTRADA)
        df_temp = pd.read_excel(xls, sheet_name=HOJA_EXCEL, header=None)

        header_row_index = -1
        for i, row in df_temp.iterrows():
             # Buscamos una celda que contenga 'Categoría' (insensible a mayúsculas/minúsculas y espacios)
            if any(str(cell).strip().lower() == 'categoría' for cell in row):
                header_row_index = i
                break

        if header_row_index == -1:
             raise ValueError("No se encontró la fila de cabecera que contiene 'Categoría'. Asegúrate de que exista.")

        # Leer nuevamente especificando la fila de cabecera correcta
        df = pd.read_excel(ARCHIVO_EXCEL_ENTRADA, sheet_name=HOJA_EXCEL, header=header_row_index)
        df = df.dropna(how='all', axis=1) # Eliminar columnas completamente vacías si las hubiera

        # Validar columnas requeridas (versión simplificada)
        print("🔍 Columnas encontradas en el Excel:", df.columns.tolist())
        for col in cols_requeridas:
            if col not in df.columns:
                print(f"⚠️ Advertencia: Columna '{col}' no encontrada. Algunas funciones podrían fallar.")

        # Renombrar columnas si es necesario para asegurar consistencia interna del script
        # (Este paso asume que las columnas existen pero podrían tener ligeras variaciones, ajustar si es necesario)
        # Ejemplo: df.rename(columns={'Emisiones 2022': 'Emisiones 2022 Real (tCO₂e)'}, inplace=True)

        # Asegurar tipos de datos numéricos donde sea necesario (ejemplo)
        numeric_cols = ['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)', 'Variación (%)', 'Indicadores por pieza 2022 (kgCO₂e)']
        for col in numeric_cols:
            if col in df.columns:
                 # Intentar convertir a numérico, los errores se convierten en NaN (Not a Number)
                 # Se eliminan símbolos como '%' antes de convertir si aplica
                 if df[col].dtype == 'object': # Solo si es texto
                     df[col] = df[col].astype(str).str.replace('%', '', regex=False) # Quitar % si existe para Variación
                 df[col] = pd.to_numeric(df[col], errors='coerce')

        # Manejar filas que son separadores o completamente vacías después de la cabecera
        df = df.dropna(subset=[cols_requeridas[0]], how='all') # Eliminar filas donde la primera columna requerida es NaN

        return df

    except FileNotFoundError:
        print(f"🚨 Error: No se encontró el archivo Excel en '{ARCHIVO_EXCEL_ENTRADA}'. Verifica la ruta.")
        sys.exit(1)
    except ValueError as ve:
         print(f"🚨 Error de Validación: {str(ve)}")
         sys.exit(1)
    except Exception as e:
        print(f"🚨 Error inesperado al cargar datos: {str(e)}")
        sys.exit(1)

# ======================================
# Análisis Predictivo (Ajustado a 2 años)
# ======================================
def predecir_emisiones(df):
    """Modelo lineal ponderado por criticidad SBTi."""
    # Asegurarse que las columnas necesarias existen y son numéricas
    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    for col in required_cols:
        if col not in df.columns:
            print(f"🚨 Error en predecir_emisiones: Falta la columna '{col}'.")
            return pd.DataFrame() # Devolver DataFrame vacío en caso de error

    # Crear copia para evitar SettingWithCopyWarning
    df_clean = df.dropna(subset=['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']).copy()

    # Filtrar filas no deseadas como subtotales o totales ANTES del modelo
    df_clean = df_clean[~df_clean['Categoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)]

    # Revisar si quedan datos después de limpiar NaNs y filtros
    if df_clean.empty:
        print("🚨 Error en predecir_emisiones: No hay datos válidos para entrenar el modelo después de la limpieza.")
        return pd.DataFrame()

    # Asignar pesos: Alcance 1 y transporte son prioritarios
    df_clean['peso'] = np.where(
        df_clean['Categoría'].astype(str).str.contains('Alcance 1|transporte', case=False, na=False),
        2.0,
        1.0
    )

    # Verificar que hay datos suficientes para entrenar
    if len(df_clean) < 2: # Se necesita al menos 2 puntos para una regresión lineal simple
         print("🚨 Error en predecir_emisiones: No hay suficientes datos para entrenar el modelo (se necesita al menos 2 filas).")
         return pd.DataFrame()

    # Entrenar modelo
    X = df_clean[['Emisiones 2021 (tCO₂e)']]
    y = df_clean['Emisiones 2022 Real (tCO₂e)']
    model = LinearRegression()
    model.fit(X, y, sample_weight=df_clean['peso'])

    # Proyectar 2023 con mejora del 5%
    df_clean['Predicción 2023 (tCO₂e)'] = model.predict(X) * 0.95

    # Devolver columnas relevantes
    return df_clean[['Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)', 'Predicción 2023 (tCO₂e)']]

# ======================================
# Simulación de Escenarios SBTi
# ======================================
def simular_escenarios(df):
    """Escenarios técnicos para mitigación."""
    if 'Subcategoría' not in df.columns or 'Emisiones 2022 Real (tCO₂e)' not in df.columns:
         print("🚨 Error en simular_escenarios: Faltan las columnas 'Subcategoría' o 'Emisiones 2022 Real (tCO₂e)'.")
         return {} # Devolver diccionario vacío

    # Trabajar con una copia limpia de datos numéricos
    df_sim = df.dropna(subset=['Subcategoría', 'Emisiones 2022 Real (tCO₂e)']).copy()
    df_sim['Emisiones 2022 Real (tCO₂e)'] = pd.to_numeric(df_sim['Emisiones 2022 Real (tCO₂e)'], errors='coerce')
    df_sim = df_sim.dropna(subset=['Emisiones 2022 Real (tCO₂e)']) # Eliminar filas donde la conversión falló

    if df_sim.empty:
         print("🚨 Error en simular_escenarios: No hay datos válidos para simulación.")
         return {}

    # Definir áreas de acción clave
    escenarios = { # <<< Nombre definido como 'escenarios'
        'Electrificación Flota Terrestre': {
            # Ajustar los nombres EXACTOS según tu Excel
            'targets': ['Transporte terrestre terceros', 'Distribución y transporte (Alcance 1)', 'Distribucion upstream'],
            'reducción': 0.45  # 45% menos emisiones
        },
        'Energía 100% Renovable': {
            # Ajustar el nombre EXACTO según tu Excel
            'targets': ['Compra de electricidad (Alcance 2)'],
            'reducción': 1.0  # Cero emisiones
        }
        # Puedes añadir más escenarios aquí
    }

    # Calcular impacto
    resultados = {}
    # Emisión base total para comparar (solo de las filas usadas en simulación)
    emision_base_2022 = df_sim['Emisiones 2022 Real (tCO₂e)'].sum()

    # *** CORRECCIÓN AQUÍ ***
    for nombre, params in escenarios.items(): # <<< Usar 'escenarios' en lugar de 'scenarios'
        # Crear columna temporal para emisiones optimizadas en esta simulación
        df_sim['Emisiones Optimizadas Temp'] = df_sim['Emisiones 2022 Real (tCO₂e)']

        # Identificar las filas que coinciden con los 'targets' del escenario actual
        # Usar .astype(str).str.contains para búsqueda flexible si los nombres no son exactos
        # Asegurarse que 'targets' sea una lista no vacía antes de unir con '|'
        if params['targets']:
             # Construir el patrón de búsqueda de forma segura
             pattern = '|'.join([re.escape(t) for t in params['targets']]) # Escapar caracteres especiales por si acaso
             mask = df_sim['Subcategoría'].astype(str).str.contains(pattern, case=False, na=False, regex=True) # Activar regex=True
        else:
             mask = pd.Series([False] * len(df_sim)) # Si no hay targets, la máscara es toda False


        # Aplicar la reducción SOLO a las filas identificadas
        df_sim.loc[mask, 'Emisiones Optimizadas Temp'] = df_sim.loc[mask, 'Emisiones 2022 Real (tCO₂e)'] * (1 - params['reducción'])

        # Sumar TODAS las emisiones (optimizadas y no optimizadas) para obtener el total del escenario
        total_optimizado = df_sim['Emisiones Optimizadas Temp'].sum()
        resultados[nombre] = round(total_optimizado, 2)

        # Opcional: Calcular la reducción lograda por este escenario
        # reduccion_escenario = emision_base_2022 - total_optimizado
        # print(f"  - Reducción estimada para {nombre}: {reduccion_escenario:,.2f} tCO₂e")

    return resultados

# ======================================
# Cumplimiento SBTi
# ======================================
def evaluar_sbti(df):
    """Verificación de métricas clave."""
    required_cols = ['Categoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    for col in required_cols:
        if col not in df.columns:
             print(f"🚨 Error en evaluar_sbti: Falta la columna '{col}'.")
             return {}

    # Asegurar que los datos son numéricos, convertir errores a NaN
    for col in ['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Requerimientos SBTi
    umbrales = {
        'reducción_anual': 0.042,  # 4.2% anual
        'cobertura_scope3': 0.67    # 67% del total Scope 3 debe ser cubierto si es significativo
    }

    # Cálculos con manejo de errores si no encuentra las filas TOTAL o Subtotal
    try:
        # Buscar la fila TOTAL (insensible a mayúsculas/minúsculas y espacios)
        total_row = df[df['Categoría'].astype(str).str.strip().str.lower() == 'total emisiones']
        if total_row.empty:
             raise ValueError("No se encontró la fila 'TOTAL EMISIONES' en la columna 'Categoría'.")
        total_2021 = total_row['Emisiones 2021 (tCO₂e)'].iloc[0]
        total_2022 = total_row['Emisiones 2022 Real (tCO₂e)'].iloc[0]

        # Buscar la fila Subtotal Alcance 3
        scope3_row = df[df['Categoría'].astype(str).str.strip().str.lower() == 'subtotal alcance 3']
        if scope3_row.empty:
             raise ValueError("No se encontró la fila 'Subtotal Alcance 3' en la columna 'Categoría'.")
        scope3_2022 = scope3_row['Emisiones 2022 Real (tCO₂e)'].iloc[0]

        # Validar que los totales no sean NaN o cero antes de dividir
        if pd.isna(total_2021) or pd.isna(total_2022) or total_2021 == 0:
            reducción_real = np.nan # No se puede calcular
        else:
            reducción_real = (total_2021 - total_2022) / total_2021

        if pd.isna(scope3_2022) or pd.isna(total_2022) or total_2022 == 0:
            cobertura_scope3_calc = np.nan # No se puede calcular
        else:
            cobertura_scope3_calc = scope3_2022 / total_2022

        # Calcular brecha y cumplimiento solo si los cálculos base son válidos
        brecha_reduccion = reducción_real - umbrales['reducción_anual'] if not pd.isna(reducción_real) else np.nan
        cumple_scope3_calc = cobertura_scope3_calc >= umbrales['cobertura_scope3'] if not pd.isna(cobertura_scope3_calc) else False
        cumple_reduccion = brecha_reduccion >= 0 if not pd.isna(brecha_reduccion) else False
        cumple_total = cumple_reduccion and cumple_scope3_calc

        return {
            'Reducción Anual Real (%)': round(reducción_real * 100, 1) if not pd.isna(reducción_real) else 'N/A',
            'Meta Reducción Anual SBTi (%)': umbrales['reducción_anual'] * 100,
            'Brecha Reducción (%)': round(brecha_reduccion * 100, 1) if not pd.isna(brecha_reduccion) else 'N/A',
            'Cobertura Alcance 3 (%)': round(cobertura_scope3_calc * 100, 1) if not pd.isna(cobertura_scope3_calc) else 'N/A',
            'Meta Cobertura Alcance 3 SBTi (%)': umbrales['cobertura_scope3'] * 100,
            'Cumple Meta Reducción': 'Sí' if cumple_reduccion else 'No',
            'Cumple Meta Cobertura Alcance 3': 'Sí' if cumple_scope3_calc else 'No',
            'Cumplimiento Global Metas Evaluadas': 'Sí' if cumple_total else 'No'
        }

    except (IndexError, ValueError, KeyError) as e:
        print(f"🚨 Error en evaluar_sbti al buscar totales o subtotales: {e}. Verifica los nombres exactos en tu Excel.")
        return {
             'Error': f"Fallo al calcular métricas SBTi: {e}"
        } # Devolver diccionario de error

# ======================================
# Visualización Profesional
# ======================================
def generar_grafico(df, ruta_guardado):
    """Gráfico de barras con contribución por categoría."""

    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
         print(f"🚨 Error en generar_grafico: Faltan columnas requeridas ({required_cols}). No se generará el gráfico.")
         return

    plt.figure(figsize=(14, 8)) # Ajuste ligero del tamaño

    # Asegurar que la columna de emisiones es numérica
    df['Emisiones 2022 Real (tCO₂e)'] = pd.to_numeric(df['Emisiones 2022 Real (tCO₂e)'], errors='coerce')

    # Filtrar solo categorías Alcance 1/2/3 válidas y excluir totales/subtotales
    df_plot = df[
        (df['Categoría'].astype(str).str.contains('Alcance [123]', case=False, na=False)) &
        (~df['Categoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)) & # Filtrar también por Categoría
        (~df['Subcategoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)) &
        (df['Emisiones 2022 Real (tCO₂e)'].notna()) & # Solo filas con valor numérico de emisión
        (df['Emisiones 2022 Real (tCO₂e)'] > 0) # Opcional: graficar solo emisiones positivas
    ].copy()

    if df_plot.empty:
        print("🎨 Advertencia en generar_grafico: No hay datos válidos para graficar después de filtrar.")
        plt.close() # Cerrar la figura vacía
        return

    # Convertir Subcategoría a string para evitar errores de tipo
    df_plot['Subcategoría'] = df_plot['Subcategoría'].astype(str)

    # Ordenar por emisiones 2022 (descendente)
    df_plot = df_plot.sort_values('Emisiones 2022 Real (tCO₂e)', ascending=False)

    # Crear mapa de colores por Alcance
    color_map = {
        '1': '#FF6B6B', # Rojo/Naranja para Alcance 1
        '2': '#4ECDC4', # Turquesa para Alcance 2
        '3': '#45B7D1'  # Azul claro para Alcance 3
    }

    # Extraer el número de alcance del string 'Categoría' para asignar color
    df_plot['Alcance_Num'] = df_plot['Categoría'].astype(str).str.extract(r'(\d)', expand=False).fillna('3') # Extrae el primer dígito, si no hay asume 3
    bar_colors = [color_map.get(alc, '#CCCCCC') for alc in df_plot['Alcance_Num']] # Usa gris si no coincide 1, 2 o 3

    # Crear gráfico
    bars = plt.bar(
        df_plot['Subcategoría'],
        df_plot['Emisiones 2022 Real (tCO₂e)'],
        color=bar_colors # Usar los colores definidos
    )

    plt.title('Contribución a las Emisiones 2022 por Subcategoría (Alcances 1, 2 y 3)', fontsize=16, pad=20) # Título más descriptivo
    plt.xlabel('Subcategoría', fontsize=12)
    plt.ylabel('Emisiones (tCO₂e)', fontsize=12) # Etiqueta Y más clara
    plt.xticks(rotation=75, ha='right', fontsize=9) # Mayor rotación para nombres largos
    plt.yticks(fontsize=10)
    plt.grid(axis='y', linestyle='--', alpha=0.6) # Rejilla más suave

    # Añadir etiquetas de valor sobre las barras
    for bar in bars:
        height = bar.get_height()
        if height > 0: # Solo etiquetar barras con altura
            plt.text(bar.get_x() + bar.get_width()/2., height, f'{height:,.0f}',
                     ha='center', va='bottom', fontsize=8.5, rotation=0) # Tamaño de fuente ligero

    # Añadir leyenda simple para los colores de Alcance
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color_map['1'], label='Alcance 1'),
                       Patch(facecolor=color_map['2'], label='Alcance 2'),
                       Patch(facecolor=color_map['3'], label='Alcance 3')]
    plt.legend(handles=legend_elements, title="Alcance", fontsize=10, title_fontsize=11)

    plt.tight_layout() # Ajustar layout para evitar solapamientos

    try:
        plt.savefig(ruta_guardado, dpi=300, bbox_inches='tight') # Guardar con buen ajuste
        print(f"✅ Gráfico guardado en: {ruta_guardado}")
    except Exception as e:
        print(f"🚨 Error al guardar el gráfico: {e}")
    finally:
        plt.close() # Asegurarse de cerrar la figura

# ======================================
# Ejecución Principal y Guardado de Resultados
# ======================================
if __name__ == "__main__":
    print("⏳ Iniciando análisis...")
    print("-" * 30)

    print("⏳ Cargando y validando datos...")
    datos_originales = cargar_datos()
    print(f"✅ Datos cargados. {len(datos_originales)} filas encontradas inicialmente.")
    print("-" * 30)

    # --- Análisis Predictivo ---
    print("🔍 Ejecutando análisis predictivo...")
    df_predicciones = predecir_emisiones(datos_originales.copy()) # Usar copia para no afectar otros análisis
    if not df_predicciones.empty:
        print(f"✅ Predicción 2023 generada para {len(df_predicciones)} subcategorías.")
        # print("\nVista previa Predicción 2023:") # Quitado para no llenar consola
        # print(df_predicciones.head().to_markdown(index=False, tablefmt="github")) # Quitado
    else:
        print("⚠️ No se pudo generar la predicción.")
    print("-" * 30)

    # --- Simulación de Escenarios ---
    print("🔍 Simulando escenarios SBTi...")
    dict_simulaciones = simular_escenarios(datos_originales.copy()) # Usar copia
    if dict_simulaciones:
        print(f"✅ {len(dict_simulaciones)} escenarios simulados.")
        # for esc, valor in dict_simulaciones.items(): # Quitado para no llenar consola
        #    print(f"  - {esc}: {valor:,.0f} tCO₂e (Emisiones Totales Estimadas)") # Quitado
    else:
        print("⚠️ No se pudieron ejecutar las simulaciones.")
    print("-" * 30)

    # --- Evaluación SBTi ---
    print("🔍 Evaluando cumplimiento SBTi...")
    dict_sbti = evaluar_sbti(datos_originales) # No necesita copia si no modifica el df
    if dict_sbti and 'Error' not in dict_sbti:
        print("✅ Evaluación SBTi completada.")
        # print(f"  - Reducción Anual: {dict_sbti.get('Reducción Anual Real (%)', 'N/A')}% (Meta: {dict_sbti.get('Meta Reducción Anual SBTi (%)', 'N/A')}%)") # Quitado
        # print(f"  - Cobertura Scope 3: {dict_sbti.get('Cobertura Alcance 3 (%)', 'N/A')}% (Meta: {dict_sbti.get('Meta Cobertura Alcance 3 SBTi (%)', 'N/A')}%)") # Quitado
        # print(f"  - Cumplimiento Global: {dict_sbti.get('Cumplimiento Global Metas Evaluadas', 'No')}") # Quitado
    elif 'Error' in dict_sbti:
        print(f"⚠️ Error en la evaluación SBTi: {dict_sbti['Error']}")
    else:
         print("⚠️ No se pudo completar la evaluación SBTi.")
    print("-" * 30)

    # --- Guardar Resultados en Excel ---
    print(f"💾 Guardando resultados en: {ARCHIVO_RESULTADOS_EXCEL}")
    try:
        with pd.ExcelWriter(ARCHIVO_RESULTADOS_EXCEL, engine='openpyxl') as writer:
            # Hoja 1: Predicciones
            if not df_predicciones.empty:
                df_predicciones.to_excel(writer, sheet_name='Predicciones_2023', index=False)
                print("  - Hoja 'Predicciones_2023' guardada.")
            else:
                 print("  - No hay datos de predicciones para guardar.")

            # Hoja 2: Simulaciones
            if dict_simulaciones:
                # Convertir diccionario de simulaciones a DataFrame para guardar
                df_simulaciones = pd.DataFrame(dict_simulaciones.items(), columns=['Escenario', 'Emisiones Optimizadas (tCO₂e)'])
                df_simulaciones.to_excel(writer, sheet_name='Simulaciones_SBTi', index=False)
                print("  - Hoja 'Simulaciones_SBTi' guardada.")
            else:
                 print("  - No hay datos de simulaciones para guardar.")

            # Hoja 3: Evaluación SBTi
            if dict_sbti and 'Error' not in dict_sbti:
                 # Convertir diccionario de evaluación a DataFrame
                 df_sbti = pd.DataFrame(dict_sbti.items(), columns=['Métrica', 'Valor'])
                 df_sbti.to_excel(writer, sheet_name='Evaluacion_SBTi', index=False)
                 print("  - Hoja 'Evaluacion_SBTi' guardada.")
            elif 'Error' in dict_sbti:
                 # Guardar el mensaje de error en la hoja
                 df_sbti_error = pd.DataFrame([{'Métrica': 'Error Evaluación', 'Valor': dict_sbti['Error']}])
                 df_sbti_error.to_excel(writer, sheet_name='Evaluacion_SBTi', index=False)
                 print("  - Hoja 'Evaluacion_SBTi' guardada (con error).")
            else:
                 print("  - No hay datos de evaluación SBTi para guardar.")

        print(f"✅ Resultados guardados exitosamente en '{ARCHIVO_RESULTADOS_EXCEL}'.")

    except Exception as e:
        print(f"🚨 Error Crítico al guardar resultados en Excel: {str(e)}")
    print("-" * 30)

    # --- Generación de Gráfico ---
    print("🎨 Generando gráfico de contribución...")
    generar_grafico(datos_originales, os.path.join(DIR_PROYECTO, 'contribucion_emisiones.png'))
    print("-" * 30) # Separador final

    print("🏁 Análisis completado.")
