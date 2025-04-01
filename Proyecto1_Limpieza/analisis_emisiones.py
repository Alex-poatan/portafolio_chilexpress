# -*- coding: utf-8 -*-
"""
Script de Análisis de Emisiones Chilexpress - Predictivo, Simulación y Cumplimiento SBTi
Autor: Asistente AI
Requerimientos: Python 3.8+, pandas, matplotlib, scikit-learn, openpyxl

ADVERTENCIAS:
- Las simulaciones asumen que los datos de entrada son correctos y en tCO₂e.
- La evaluación SBTi es una aproximación simplificada, no un certificado oficial.
- Las predicciones 2023 no consideran variables externas (ej: crecimiento operacional).
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
ARCHIVO_RESULTADOS_EXCEL = os.path.join(DIR_PROYECTO, "resultados_analisis_emisiones.xlsx")
ARCHIVO_GRAFICO = os.path.join(DIR_PROYECTO, 'contribucion_emisiones.png')

# ======================================
# Funciones de Carga y Validación (Mejoradas)
# ======================================
def cargar_datos():
    """Carga y valida estructura del Excel con conversión de unidades."""
    cols_requeridas = [
        'Categoría', 'Subcategoría', 'Emisiones 2021 (tCO₂e)',
        'Emisiones 2022 Real (tCO₂e)', 'Metodología/Factor Emisión',
        'Notas Ajustadas', 'Variación (%)', 'Indicadores por pieza 2022 (kgCO₂e)'
    ]
    try:
        xls = pd.ExcelFile(ARCHIVO_EXCEL_ENTRADA)
        df_temp = pd.read_excel(xls, sheet_name=HOJA_EXCEL, header=None)

        header_row_index = -1
        for i, row in df_temp.iterrows():
            if any(str(cell).strip().lower() == 'categoría' for cell in row):
                header_row_index = i
                break

        if header_row_index == -1:
            raise ValueError("No se encontró la fila de cabecera que contiene 'Categoría'. Asegúrate de que exista.")

        df = pd.read_excel(ARCHIVO_EXCEL_ENTRADA, sheet_name=HOJA_EXCEL, header=header_row_index)
        df = df.dropna(how='all', axis=1)

        print("🔍 Columnas encontradas en el Excel:", df.columns.tolist())
        for col in cols_requeridas:
            if col not in df.columns:
                print(f"⚠️ Advertencia: Columna '{col}' no encontrada. Algunas funciones podrían fallar.")

        # Conversión automática de unidades si se detectan valores altos (kg -> t)
        numeric_cols = ['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
        for col in numeric_cols:
            if col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).str.replace('%', '', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce')
                # Si valores > 1 millón, convertir kg a toneladas
                if not df[col].dropna().empty and df[col].max() > 1e6:
                    print(f"⚠️ CONVERSIÓN: Valores en '{col}' convertidos de kg a tCO₂e (divididos por 1000).")
                    df[col] = df[col] / 1000

        df = df.dropna(subset=[cols_requeridas[0]], how='all')
        return df  # <- Asegurar return

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
# Análisis Predictivo (Corregido y Mejorado)
# ======================================
def predecir_emisiones(df):
    """Modelo lineal usando 2021 para predecir 2022, luego proyecta 2023."""
    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
         print(f"🚨 Error en predecir_emisiones: Faltan columnas requeridas ({required_cols}).")
         return pd.DataFrame()

    df_clean = df.dropna(subset=['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']).copy()
    df_clean = df_clean[~df_clean['Categoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)]

    if df_clean.empty or len(df_clean) < 2:
        print("🚨 Error en predecir_emisiones: No hay suficientes datos válidos para entrenar el modelo.")
        return pd.DataFrame()

    # Entrenar modelo con 2021 -> 2022
    X = df_clean[['Emisiones 2021 (tCO₂e)']].copy()  # Usar copia explícita
    y = df_clean['Emisiones 2022 Real (tCO₂e)'].copy()

    model = LinearRegression()
    model.fit(X, y)

    # Proyectar 2023 usando 2022 como base (con nombres de características consistentes)
    X_2022 = df_clean[['Emisiones 2022 Real (tCO₂e)']].rename(columns={'Emisiones 2022 Real (tCO₂e)': 'Emisiones 2021 (tCO₂e)'})
    df_clean['Predicción 2023 (tCO₂e)'] = model.predict(X_2022) * 0.95

    return df_clean[['Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)', 'Predicción 2023 (tCO₂e)']]

# ======================================
# Simulación de Escenarios SBTi (Corregida)
# ======================================
def simular_escenarios(df):
    """Escenarios técnicos para mitigación."""
    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
        print(f"🚨 Error en simular_escenarios: Faltan columnas requeridas ({required_cols}).")
        return {}

    df_sim = df.dropna(subset=['Subcategoría', 'Emisiones 2022 Real (tCO₂e)']).copy()
    df_sim['Emisiones 2022 Real (tCO₂e)'] = pd.to_numeric(df_sim['Emisiones 2022 Real (tCO₂e)'], errors='coerce')
    df_sim = df_sim.dropna(subset=['Emisiones 2022 Real (tCO₂e)'])

    # Limpiar espacios sobrantes en "Subcategoría"
    df_sim['Subcategoría'] = df_sim['Subcategoría'].astype(str).str.strip()

    # Imprimir subcategorías únicas para ver qué se está leyendo
    print("Subcategorías únicas encontradas:")
    print(df_sim['Subcategoría'].unique())

    # Filtrar filas no deseadas ANTES de calcular la base y aplicar escenarios
    df_sim = df_sim[
        ~df_sim['Categoría'].astype(str).str.contains('Subtotal|TOTAL', case=False, na=False) &
        ~df_sim['Subcategoría'].astype(str).str.contains('Subtotal|TOTAL|encomiendas', case=False, na=False)
    ].copy()

    if df_sim.empty:
        print("🚨 Error en simular_escenarios: No hay datos válidos para simulación después de filtrar.")
        return {}

    # Emisión base total para comparar (solo de las filas válidas para simulación)
    emision_base_2022 = df_sim['Emisiones 2022 Real (tCO₂e)'].sum()
    print(f"ℹ️ Emisión base 2022 (Alcances 1, 2, 3 - filtrado) para simulación: {emision_base_2022:,.2f} tCO₂e")

    # Actualizar targets con los valores existentes en df_sim
    escenarios = {
        'Electrificación Flota Terrestre': {
            'targets': [
                'Combustión - fuentes fijas'
            ],
            'reducción': 0.45
        },
        'Energía 100% Renovable': {
            'targets': [
                'Compra de electricidad'
            ],
            'reducción': 1.0
        }
    }

    resultados = {}
    for nombre, params in escenarios.items():
        target_mask = df_sim['Subcategoría'].isin(params['targets'])
        if not target_mask.any():
             print(f"⚠️ Advertencia Simulación: No se encontraron subcategorías válidas {params['targets']} para el escenario '{nombre}'. Verifica los nombres en `escenarios` y en tu Excel.")
             continue

        df_sim['Emisiones Optimizadas Temp'] = df_sim['Emisiones 2022 Real (tCO₂e)']
        df_sim.loc[target_mask, 'Emisiones Optimizadas Temp'] = df_sim.loc[target_mask, 'Emisiones 2022 Real (tCO₂e)'] * (1 - params['reducción'])
        total_optimizado = pd.to_numeric(df_sim['Emisiones Optimizadas Temp'], errors='coerce').sum()

        if pd.isna(total_optimizado):
             print(f"🚨 Error calculando el total optimizado para el escenario '{nombre}'.")
             resultados[nombre] = 'Error de Cálculo'
        else:
             resultados[nombre] = round(total_optimizado, 2)
             reduccion_lograda = emision_base_2022 - total_optimizado
             print(f"   - Escenario '{nombre}': Emisiones Totales Estimadas = {total_optimizado:,.2f} tCO₂e (Reducción vs base: {reduccion_lograda:,.2f} tCO₂e)")

    return resultados

# ======================================
# Cumplimiento SBTi
# ======================================
def evaluar_sbti(df):
    """Verificación de métricas clave."""
    required_cols = ['Categoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
            print(f"🚨 Error en evaluar_sbti: Faltan columnas requeridas ({required_cols}).")
            return {'Error': f"Faltan columnas: {required_cols}"}

    for col in ['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']:
         if col in df.columns:
             df[col] = pd.to_numeric(df[col], errors='coerce')
         else:
              print(f"🚨 Error Crítico en evaluar_sbti: La columna requerida '{col}' no existe en los datos cargados.")
              return {'Error': f"Columna faltante: {col}"}

    umbrales = {
        'reducción_anual': 0.042,  # 4.2% anual (ejemplo WB2C)
        'cobertura_scope3': 0.67   # 67% del total Scope 3 debe ser cubierto si es significativo (>40% total)
    }

    try:
        total_row = df[df['Categoría'].astype(str).str.strip().str.lower() == 'total emisiones']
        scope3_row = df[df['Categoría'].astype(str).str.strip().str.lower() == 'subtotal alcance 3']

        if total_row.empty: raise ValueError("Fila 'TOTAL EMISIONES' no encontrada.")
        if scope3_row.empty: raise ValueError("Fila 'Subtotal Alcance 3' no encontrada.")

        total_2021 = total_row['Emisiones 2021 (tCO₂e)'].iloc[0]
        total_2022 = total_row['Emisiones 2022 Real (tCO₂e)'].iloc[0]
        scope3_2022 = scope3_row['Emisiones 2022 Real (tCO₂e)'].iloc[0]

        if pd.isna(total_2021) or pd.isna(total_2022) or pd.isna(scope3_2022):
             raise ValueError("Valores NaN encontrados en Totales o Subtotal Alcance 3.")
        if total_2021 == 0 or total_2022 == 0:
             print("⚠️ Advertencia SBTi: Emisiones totales 2021 o 2022 son cero. Algunos cálculos pueden ser N/A.")
             reducción_real = np.nan if total_2021 == 0 else (total_2021 - total_2022) / total_2021
             cobertura_scope3_calc = np.nan if total_2022 == 0 else scope3_2022 / total_2022
        else:
             reducción_real = (total_2021 - total_2022) / total_2021
             cobertura_scope3_calc = scope3_2022 / total_2022

        scope3_es_significativo = cobertura_scope3_calc >= 0.40 if not pd.isna(cobertura_scope3_calc) else False
        brecha_reduccion = reducción_real - umbrales['reducción_anual'] if not pd.isna(reducción_real) else np.nan
        cumple_scope3_calc = scope3_es_significativo and (cobertura_scope3_calc >= umbrales['cobertura_scope3'] if not pd.isna(cobertura_scope3_calc) else False)
        cumple_reduccion = brecha_reduccion >= 0 if not pd.isna(brecha_reduccion) else False
        cumple_total = cumple_reduccion and cumple_scope3_calc

        return {
            'Reducción Anual Real (%)': round(reducción_real * 100, 1) if not pd.isna(reducción_real) else 'N/A',
            'Meta Reducción Anual SBTi (%)': umbrales['reducción_anual'] * 100,
            'Brecha Reducción (%)': round(brecha_reduccion * 100, 1) if not pd.isna(brecha_reduccion) else 'N/A',
            'Cobertura Alcance 3 (% del Total)': round(cobertura_scope3_calc * 100, 1) if not pd.isna(cobertura_scope3_calc) else 'N/A',
            'Alcance 3 Significativo (>40%)': 'Sí' if scope3_es_significativo else 'No',
            'Meta Cobertura Alcance 3 SBTi (%) (si aplica)': umbrales['cobertura_scope3'] * 100 if scope3_es_significativo else 'N/A',
            'Cumple Meta Reducción': 'Sí' if cumple_reduccion else 'No',
            'Cumple Meta Cobertura Alcance 3': 'Sí' if cumple_scope3_calc else 'No',
            'Cumplimiento Global Metas Evaluadas': 'Sí' if cumple_total else 'No'
        }

    except (IndexError, ValueError, KeyError) as e:
        print(f"🚨 Error en evaluar_sbti: {e}. Verifica nombres exactos y valores numéricos en filas TOTAL/Subtotal de tu Excel.")
        return {'Error': f"Fallo al calcular métricas SBTi: {e}"}

# ======================================
# Visualización Profesional (Mejorada)
# ======================================
def generar_grafico(df, ruta_guardado):
    """Gráfico de barras con contribución por categoría."""
    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
        print(f"🚨 Error en generar_grafico: Faltan columnas requeridas ({required_cols}). No se generará el gráfico.")
        return

    plt.figure(figsize=(14, 8))
    df_plot = df.copy()
    df_plot['Emisiones 2022 Real (tCO₂e)'] = pd.to_numeric(df_plot['Emisiones 2022 Real (tCO₂e)'], errors='coerce')

    df_plot = df_plot[
        (df_plot['Categoría'].astype(str).str.contains('Alcance [123]', case=False, na=False)) &
        (~df_plot['Categoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)) &
        (~df_plot['Subcategoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)) &
        (df_plot['Emisiones 2022 Real (tCO₂e)'].notna()) &
        (df_plot['Emisiones 2022 Real (tCO₂e)'] > 0)
    ].copy()

    if df_plot.empty:
        print("🎨 Advertencia en generar_grafico: No hay datos válidos para graficar después de filtrar.")
        plt.close()
        return

    # Asegurar conversión de unidades en el gráfico:
    df_plot['Emisiones 2022 Real (tCO₂e)'] = df_plot['Emisiones 2022 Real (tCO₂e)'] / 1000

    df_plot['Subcategoría'] = df_plot['Subcategoría'].astype(str)
    df_plot = df_plot.sort_values('Emisiones 2022 Real (tCO₂e)', ascending=False)

    color_map = {'1': '#FF6B6B', '2': '#4ECDC4', '3': '#45B7D1'}
    df_plot['Alcance_Num'] = df_plot['Categoría'].astype(str).str.extract(r'(\d)', expand=False).fillna('3')
    bar_colors = [color_map.get(alc, '#CCCCCC') for alc in df_plot['Alcance_Num']]

    bars = plt.bar(
        df_plot['Subcategoría'], df_plot['Emisiones 2022 Real (tCO₂e)'], color=bar_colors
    )

    plt.title('Contribución a las Emisiones 2022 por Subcategoría (Alcances 1, 2 y 3)', fontsize=16, pad=20)
    plt.xlabel('Subcategoría', fontsize=12)
    plt.ylabel('Emisiones (tCO₂e)', fontsize=12)
    plt.xticks(rotation=75, ha='right', fontsize=9)
    plt.yticks(fontsize=10)
    plt.grid(axis='y', linestyle='--', alpha=0.6)

    for bar in bars:
        height = bar.get_height()
        if height > 0:
            plt.text(bar.get_x() + bar.get_width()/2., height, f'{height:,.0f}',
                     ha='center', va='bottom', fontsize=8.5, rotation=0)

    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color_map['1'], label='Alcance 1'),
                       Patch(facecolor=color_map['2'], label='Alcance 2'),
                       Patch(facecolor=color_map['3'], label='Alcance 3')]
    plt.legend(handles=legend_elements, title="Alcance", fontsize=10, title_fontsize=11)

    plt.tight_layout()
    try:
        plt.savefig(ruta_guardado, dpi=300, bbox_inches='tight')
        print(f"✅ Gráfico guardado en: {ruta_guardado}")
    except Exception as e:
        print(f"🚨 Error al guardar el gráfico: {e}")
    finally:
        plt.close()

# ======================================
# Generación de Reporte en Texto
# ======================================
def generar_reporte_texto(dict_sbti, dict_simulaciones, df_predicciones):
    """Genera un resumen ejecutivo en texto formateado."""
    reporte = "="*50 + "\n"
    reporte += "        RESUMEN EJECUTIVO ANÁLISIS DE EMISIONES\n"
    reporte += "="*50 + "\n\n"

    reporte += "--- Evaluación SBTi Simplificada (Basada en 2021-2022) ---\n"
    if dict_sbti and 'Error' not in dict_sbti:
        reporte += f"- Reducción Anual Real (2021-2022):      {dict_sbti.get('Reducción Anual Real (%)', 'N/A')} %\n"
        reporte += f"- Meta Reducción Anual (Ejemplo SBTi): {dict_sbti.get('Meta Reducción Anual SBTi (%)', 'N/A')} %\n"
        reporte += f"- Cumple Meta Reducción:                {dict_sbti.get('Cumple Meta Reducción', 'N/A')}\n\n"
        reporte += f"- Cobertura Alcance 3 (% del Total):    {dict_sbti.get('Cobertura Alcance 3 (% del Total)', 'N/A')} %\n"
        reporte += f"- Alcance 3 Significativo (>40%):       {dict_sbti.get('Alcance 3 Significativo (>40%)', 'N/A')}\n"
        reporte += f"- Meta Cobertura Alcance 3 SBTi (%) (si aplica): {dict_sbti.get('Meta Cobertura Alcance 3 SBTi (%) (si aplica)', 'N/A')} %\n"
        reporte += f"- Cumple Meta Cobertura Alcance 3:    {dict_sbti.get('Cumple Meta Cobertura Alcance 3', 'N/A')}\n\n"
        reporte += f"- CUMPLIMIENTO GLOBAL (Metas Simples):  {dict_sbti.get('Cumplimiento Global Metas Evaluadas', 'N/A')}\n"
    elif 'Error' in dict_sbti:
         reporte += f"ERROR en la evaluación: {dict_sbti['Error']}\n"
    else:
        reporte += "Evaluación SBTi no disponible.\n"
    reporte += "-"*50 + "\n\n"

    reporte += "--- Simulación de Escenarios (Impacto en tCO₂e Totales 2022) ---\n"
    if dict_simulaciones:
        for esc, valor in dict_simulaciones.items():
             if isinstance(valor, (int, float)):
                 reporte += f"- {esc}: {valor:,.2f} tCO₂e (Emisiones Totales Estimadas)\n"
             else:
                  reporte += f"- {esc}: {valor}\n"
    else:
        reporte += "Simulaciones no disponibles o no ejecutadas.\n"
    reporte += "-"*50 + "\n\n"

    reporte += "--- Predicción Simplificada 2023 (Basada en tendencia 21-22 y -5%) ---\n"
    if not df_predicciones.empty:
        top_pred = df_predicciones.nlargest(3, 'Predicción 2023 (tCO₂e)')
        for index, row in top_pred.iterrows():
             reporte += f"- {row['Subcategoría']}: {row['Predicción 2023 (tCO₂e)']:,.2f} tCO₂e\n"
        total_predicho = df_predicciones['Predicción 2023 (tCO₂e)'].sum()
        reporte += f"\n- TOTAL Predicho (suma subcat.): {total_predicho:,.2f} tCO₂e\n"
    else:
        reporte += "Predicciones no disponibles.\n"
    reporte += "="*50 + "\n"

    return reporte

# ======================================
# Ejecución Principal (Sin Cambios)
# ======================================
if __name__ == "__main__":
    print("⏳ Iniciando análisis...")
    print("-" * 30)

    datos_originales = cargar_datos()
    if datos_originales is None:
         sys.exit(1)
    print(f"✅ Datos cargados. {len(datos_originales)} filas encontradas inicialmente.")
    print("-" * 30)

    df_predicciones = predecir_emisiones(datos_originales.copy())
    if not df_predicciones.empty:
        print(f"✅ Predicción 2023 generada para {len(df_predicciones)} subcategorías.")
    else:
        print("⚠️ No se pudo generar la predicción.")
    print("-" * 30)

    dict_simulaciones = simular_escenarios(datos_originales.copy())
    if dict_simulaciones:
        print(f"✅ {len(dict_simulaciones)} escenarios simulados.")
    else:
        print("⚠️ No se pudieron ejecutar las simulaciones o no se encontraron targets.")
    print("-" * 30)

    dict_sbti = evaluar_sbti(datos_originales.copy())
    if dict_sbti and 'Error' not in dict_sbti:
        print("✅ Evaluación SBTi (simplificada) completada.")
    elif 'Error' in dict_sbti:
        print(f"⚠️ Error en la evaluación SBTi: {dict_sbti['Error']}")
    else:
        print("⚠️ No se pudo completar la evaluación SBTi.")
    print("-" * 30)

    print(f"💾 Guardando resultados tabulares en: {ARCHIVO_RESULTADOS_EXCEL}")
    try:
        with pd.ExcelWriter(ARCHIVO_RESULTADOS_EXCEL, engine='openpyxl') as writer:
            if not df_predicciones.empty:
                df_predicciones.to_excel(writer, sheet_name='Predicciones_2023', index=False)
                print("   - Hoja 'Predicciones_2023' guardada.")
            else:
                print("   - No hay datos de predicciones para guardar.")

            if dict_simulaciones:
                df_simulaciones = pd.DataFrame(dict_simulaciones.items(), columns=['Escenario', 'Emisiones Optimizadas (tCO₂e)'])
                df_simulaciones.to_excel(writer, sheet_name='Simulaciones_SBTi', index=False)
                print("   - Hoja 'Simulaciones_SBTi' guardada.")
            else:
                print("   - No hay datos de simulaciones para guardar.")

            if dict_sbti:
                df_sbti = pd.DataFrame(dict_sbti.items(), columns=['Métrica', 'Valor'])
                df_sbti.to_excel(writer, sheet_name='Evaluacion_SBTi', index=False)
                if 'Error' in dict_sbti:
                     print("   - Hoja 'Evaluacion_SBTi' guardada (con error).")
                else:
                     print("   - Hoja 'Evaluacion_SBTi' guardada.")
            else:
                print("   - No hay datos de evaluación SBTi para guardar.")

        print(f"✅ Resultados tabulares guardados exitosamente.")

    except Exception as e:
        print(f"🚨 Error Crítico al guardar resultados en Excel: {str(e)}")
    print("-" * 30)

    print("🎨 Generando gráfico de contribución...")
    generar_grafico(datos_originales, ARCHIVO_GRAFICO)
    print("-" * 30)

    print("📄 Generando reporte resumen en consola...")
    reporte_final = generar_reporte_texto(dict_sbti, dict_simulaciones, df_predicciones)
    print("\n" + reporte_final)
    print("-" * 30)
    print("🏁 Análisis completado.")
