# -*- coding: utf-8 -*-
"""
Script de Análisis de Emisiones Chilexpress - Predictivo, Simulación y Cumplimiento SBTi
Autor: Alex
Requerimientos: Python 3.8+, pandas, matplotlib, scikit-learn, openpyxl

ADVERTENCIAS:
- ESTE SCRIPT UTILIZA DATOS SIMULADOS Y ESTIMADOS.
- Las simulaciones asumen que los datos de entrada (ahora simulados) son correctos y en tCO₂e.
- La evaluación SBTi es una aproximación simplificada, no un certificado oficial.
- Las predicciones 2023 no consideran variables externas (ej: crecimiento operacional real).
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
import matplotlib.font_manager as fm

plt.rcParams['font.family'] = 'Times New Roman'

# Configurar rutas específicas (Mantener por si se usan en el futuro)
DIR_PROYECTO = r"C:\Users\Alexander\OneDrive\Escritorio\portafolio_chilexpress\Proyecto1_Limpieza" # Ajusta si es necesario
# ARCHIVO_EXCEL_ENTRADA = os.path.join(DIR_PROYECTO, "emisiones_chilexpress.xlsx") # Ya no se usará para cargar datos principales
# HOJA_EXCEL = "hoja1"
ARCHIVO_RESULTADOS_EXCEL = os.path.join(DIR_PROYECTO, "resultados_analisis_emisiones_REVISADO.xlsx") # Nuevo nombre para no sobrescribir
ARCHIVO_GRAFICO = os.path.join(DIR_PROYECTO, 'contribucion_emisiones_REVISADO.png') # Nuevo nombre para no sobrescribir

# ======================================
# *** MODIFICACIÓN: Crear Datos Simulados Representativos ***
# ======================================
def crear_datos_simulados():
    """
    Crea un DataFrame con datos de emisión simulados y más representativos
    para Chilexpress, enfocándose en las categorías clave mencionadas.
    Estos valores son ESTIMACIONES y NO datos reales.
    """
    print("⚠️  Usando DATOS SIMULADOS representativos en lugar de cargar desde Excel.")
    data = {
        'Categoría': [
            'Alcance 1',
            'Alcance 2',
            'Alcance 3',
            # Filas requeridas por la función evaluar_sbti
            'Subtotal Alcance 3',
            'TOTAL EMISIONES'
        ],
        'Subcategoría': [
            'Combustión - fuentes fijas',
            'Compra de electricidad',
            'Otras emisiones energía (vapor)',
            # Nombres correspondientes para las filas de totales
            'Subtotal Alcance 3',
            'TOTAL EMISIONES'
        ],
        'Emisiones 2021 (tCO₂e)': [
            2100,  # Estimación para Combustión fija 2021
            9500,  # Estimación para Electricidad 2021
            550,   # Estimación para Vapor (Scope 3) 2021
            # --- Totales Calculados ---
            550,   # Suma de Scope 3 (solo 'Vapor' en este ejemplo) para 2021
            12150  # Suma Total (2100 + 9500 + 550) para 2021
        ],
        'Emisiones 2022 Real (tCO₂e)': [
            2000,  # Estimación para Combustión fija 2022
            10000, # Estimación para Electricidad 2022
            500,   # Estimación para Vapor (Scope 3) 2022
            # --- Totales Calculados ---
            500,   # Suma de Scope 3 para 2022
            12500  # Suma Total (2000 + 10000 + 500) para 2022
        ],
        # Se pueden añadir valores dummy o NaN para otras columnas si son necesarias
        # en alguna función, aunque las principales (predicción, simulación, sbti, gráfico)
        # se centran en las columnas de emisión, categoría y subcategoría.
        'Metodología/Factor Emisión': ['Estimado']*5,
        'Notas Ajustadas': ['Datos simulados']*5,
        'Variación (%)': [np.nan]*5,
        'Indicadores por pieza 2022 (kgCO₂e)': [np.nan]*5
    }
    df = pd.DataFrame(data)

    # Asegurar tipos numéricos correctos
    numeric_cols = ['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Validar estructura básica
    cols_requeridas_minimas = ['Categoría', 'Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    for col in cols_requeridas_minimas:
        if col not in df.columns:
             raise ValueError(f"Error crítico: La columna simulada '{col}' no se generó correctamente.")
        if df[col].isnull().any() and col in numeric_cols : # Chequeo básico de nulos en columnas numéricas clave
             print(f"Advertencia: La columna simulada '{col}' contiene valores nulos inesperados.")

    return df

# ======================================
# Funciones de Carga y Validación (Original - Ahora no se usa para datos principales)
# ======================================
# def cargar_datos():
#     """Carga y valida estructura del Excel con conversión de unidades."""
#     # ... (código original de carga de Excel omitido ya que usamos datos simulados)
#     # ... (se mantiene por si se quiere revertir al uso de Excel en el futuro)
#     pass # Dejar vacío o comentado si no se usa

# ======================================
# Análisis Predictivo (Sin cambios en lógica, usará nuevos datos)
# ======================================
def predecir_emisiones(df):
    """Modelo lineal usando 2021 para predecir 2022, luego proyecta 2023."""
    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
          print(f"🚨 Error en predecir_emisiones: Faltan columnas requeridas ({required_cols}).")
          return pd.DataFrame()

    # Excluir filas de totales/subtotales ANTES de entrenar
    df_clean = df[
        ~df['Categoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False) &
        ~df['Subcategoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)
    ].copy()

    df_clean = df_clean.dropna(subset=['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)'])


    if df_clean.empty or len(df_clean) < 2:
        # Con pocos datos (3 categorías), la regresión lineal puede no ser muy robusta,
        # pero mantenemos la estructura. Una alternativa sería un % de cambio promedio.
        print("⚠️ Advertencia en predecir_emisiones: Pocos datos para entrenar modelo lineal robusto.")
        # Alternativa simple si hay pocos datos: Proyectar cambio promedio o fijo
        if not df_clean.empty:
             df_clean['Var_21_22'] = (df_clean['Emisiones 2022 Real (tCO₂e)'] - df_clean['Emisiones 2021 (tCO₂e)']) / df_clean['Emisiones 2021 (tCO₂e)']
             cambio_promedio = df_clean['Var_21_22'].mean()
             print(f"   Usando cambio promedio ({cambio_promedio:.2%}) para predicción 2023.")
             # Aplicar reducción adicional del 5% como en el original
             df_clean['Predicción 2023 (tCO₂e)'] = df_clean['Emisiones 2022 Real (tCO₂e)'] * (1 + cambio_promedio) * 0.95
             return df_clean[['Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)', 'Predicción 2023 (tCO₂e)']]
        else:
             print("🚨 Error en predecir_emisiones: No hay suficientes datos válidos para entrenar.")
             return pd.DataFrame()


    # Entrenar modelo con 2021 -> 2022 (si hay suficientes datos)
    X = df_clean[['Emisiones 2021 (tCO₂e)']].copy()
    y = df_clean['Emisiones 2022 Real (tCO₂e)'].copy()

    model = LinearRegression()
    model.fit(X, y)

    # Proyectar 2023 usando 2022 como base (con nombres de características consistentes)
    # Usar el modelo entrenado para proyectar desde 2022. El coeficiente indicará la tendencia.
    X_2022 = df_clean[['Emisiones 2022 Real (tCO₂e)']].rename(columns={'Emisiones 2022 Real (tCO₂e)': 'Emisiones 2021 (tCO₂e)'})
    # Aplicar reducción adicional del 5% como en el original
    df_clean['Predicción 2023 (tCO₂e)'] = model.predict(X_2022) * 0.95
    df_clean['Predicción 2023 (tCO₂e)'] = df_clean['Predicción 2023 (tCO₂e)'].clip(lower=0) # Evitar predicciones negativas

    return df_clean[['Subcategoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)', 'Predicción 2023 (tCO₂e)']]

# ======================================
# Simulación de Escenarios SBTi (Sin cambios en lógica, usará nuevos datos)
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

    df_sim['Subcategoría'] = df_sim['Subcategoría'].astype(str).str.strip()

    # Filtrar filas no deseadas ANTES de calcular la base y aplicar escenarios
    df_sim_filt = df_sim[
        ~df_sim['Categoría'].astype(str).str.contains('Subtotal|TOTAL', case=False, na=False) &
        ~df_sim['Subcategoría'].astype(str).str.contains('Subtotal|TOTAL|encomiendas', case=False, na=False) # Mantener filtro original por si acaso
    ].copy()

    if df_sim_filt.empty:
        print("🚨 Error en simular_escenarios: No hay datos válidos para simulación después de filtrar.")
        return {}

    # Emisión base total para comparar (solo de las filas válidas para simulación)
    emision_base_2022_simulacion = df_sim_filt['Emisiones 2022 Real (tCO₂e)'].sum()
    print(f"ℹ️ Emisión base 2022 (Subcategorías simulables) para simulación: {emision_base_2022_simulacion:,.2f} tCO₂e")

    # Escenarios definidos (ajustar targets si los nombres de subcategoría cambiaran)
    escenarios = {
        'Electrificación Fuentes Fijas (Ej: Calderas)': { # Nombre ajustado para claridad
            'targets': [
                'Combustión - fuentes fijas' # Asegúrate que este nombre coincide EXACTO con 'Subcategoría'
            ],
            'reducción': 0.45 # Reducción del 45% para este escenario
        },
        'Energía 100% Renovable (Electricidad)': { # Nombre ajustado para claridad
            'targets': [
                'Compra de electricidad' # Asegúrate que este nombre coincide EXACTO con 'Subcategoría'
            ],
            'reducción': 1.0 # Reducción del 100% (ideal) para este escenario
        }
        # Se podrían añadir más escenarios si hubiera más categorías (ej: flota)
    }

    resultados = {}
    df_resultados_detalle = df_sim_filt.copy() # Copia para no modificar original dentro del loop

    for nombre, params in escenarios.items():
        df_temp = df_sim_filt.copy() # Trabajar sobre una copia para cada escenario independiente
        target_mask = df_temp['Subcategoría'].isin(params['targets'])

        if not target_mask.any():
            print(f"⚠️ Advertencia Simulación: No se encontraron subcategorías válidas {params['targets']} para el escenario '{nombre}'. Verifica los nombres.")
            resultados[nombre] = 'Target no encontrado'
            continue

        # Aplicar la reducción solo a las filas target
        df_temp.loc[target_mask, 'Emisiones Optimizadas Temp'] = df_temp.loc[target_mask, 'Emisiones 2022 Real (tCO₂e)'] * (1 - params['reducción'])
        # Las filas no target mantienen su valor original
        df_temp.loc[~target_mask, 'Emisiones Optimizadas Temp'] = df_temp.loc[~target_mask, 'Emisiones 2022 Real (tCO₂e)']

        total_optimizado = pd.to_numeric(df_temp['Emisiones Optimizadas Temp'], errors='coerce').sum()

        if pd.isna(total_optimizado):
            print(f"🚨 Error calculando el total optimizado para el escenario '{nombre}'.")
            resultados[nombre] = 'Error de Cálculo'
        else:
            resultados[nombre] = round(total_optimizado, 2)
            reduccion_lograda = emision_base_2022_simulacion - total_optimizado
            print(f"   - Escenario '{nombre}': Emisiones Totales Estimadas = {total_optimizado:,.2f} tCO₂e (Reducción vs base simulable: {reduccion_lograda:,.2f} tCO₂e)")
            # Guardar detalle si se necesita
            # df_resultados_detalle[f'Sim_{nombre}'] = df_temp['Emisiones Optimizadas Temp']


    return resultados # Devuelve el total por escenario


# ======================================
# Cumplimiento SBTi (Sin cambios en lógica, usará nuevos datos y totales)
# ======================================
def evaluar_sbti(df):
    """Verificación de métricas clave."""
    required_cols = ['Categoría', 'Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
            print(f"🚨 Error en evaluar_sbti: Faltan columnas requeridas ({required_cols}).")
            return {'Error': f"Faltan columnas: {required_cols}"}

    # Asegurarse que las columnas de emisión son numéricas
    for col in ['Emisiones 2021 (tCO₂e)', 'Emisiones 2022 Real (tCO₂e)']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        else:
             print(f"🚨 Error Crítico en evaluar_sbti: La columna requerida '{col}' no existe.")
             return {'Error': f"Columna faltante: {col}"}

    # Umbrales SBTi (Ejemplos - estos deben definirse según el target real de la empresa)
    umbrales = {
        'reducción_anual': 0.042,  # 4.2% anual (ejemplo para Well-Below 2°C)
        'cobertura_scope3': 0.67   # 67% del total Scope 3 debe ser cubierto por un target si Scope 3 es > 40% del total (Scopes 1+2+3)
    }

    try:
        # Buscar filas de totales usando los nombres exactos definidos en crear_datos_simulados
        total_row = df[df['Subcategoría'].astype(str).str.strip().str.lower() == 'total emisiones']
        scope3_row = df[df['Subcategoría'].astype(str).str.strip().str.lower() == 'subtotal alcance 3']

        if total_row.empty: raise ValueError("Fila 'TOTAL EMISIONES' no encontrada en datos simulados.")
        if scope3_row.empty: raise ValueError("Fila 'Subtotal Alcance 3' no encontrada en datos simulados.")

        # Extraer valores totales y de alcance 3
        total_2021 = total_row['Emisiones 2021 (tCO₂e)'].iloc[0]
        total_2022 = total_row['Emisiones 2022 Real (tCO₂e)'].iloc[0]
        scope3_2022 = scope3_row['Emisiones 2022 Real (tCO₂e)'].iloc[0]

        if pd.isna(total_2021) or pd.isna(total_2022) or pd.isna(scope3_2022):
            raise ValueError("Valores NaN encontrados en Totales o Subtotal Alcance 3 simulados.")

        # Calcular métricas
        if total_2021 == 0:
             print("⚠️ Advertencia SBTi: Emisiones totales 2021 son cero. Cálculo de reducción es N/A.")
             reducción_real = np.nan
        else:
             reducción_real = (total_2021 - total_2022) / total_2021 # Ojo: aquí dará negativo si 2022 > 2021

        if total_2022 == 0:
             print("⚠️ Advertencia SBTi: Emisiones totales 2022 son cero. Cálculo de cobertura Scope 3 es N/A.")
             cobertura_scope3_calc = np.nan
        else:
            # Cobertura de las emisiones de Scope 3 sobre el TOTAL (Scope 1+2+3)
             cobertura_scope3_calc = scope3_2022 / total_2022

        # Evaluar cumplimiento (Simplificado)
        scope3_es_significativo = cobertura_scope3_calc >= 0.40 if not pd.isna(cobertura_scope3_calc) else False

        # La "reducción real" calculada es la observada 21-22. El "cumplimiento" es si esta reducción observada
        # alcanza o supera la meta anual *requerida* por SBTi (ej: 4.2%).
        # Nota: Una reducción observada negativa (aumento de emisiones) obviamente no cumple.
        cumple_reduccion = (reducción_real >= umbrales['reducción_anual']) if not pd.isna(reducción_real) else False

        # Para Scope 3, SBTi requiere que *si* es significativo (>40%), entonces un target debe cubrir
        # al menos el 67% de esas emisiones de Scope 3. Esta evaluación simplificada sólo chequea si es significativo.
        # No podemos verificar si un target *cubre* el 67% sin saber qué subcategorías incluye el target de Scope 3.
        # Aquí simplificamos: si es significativo, asumimos que se requiere un target, pero no evaluamos la cobertura del target en sí.
        requiere_target_scope3 = scope3_es_significativo
        # Evaluación simplificada de cumplimiento Scope 3 podría ser si se tiene un target definido (no evaluable aquí)
        cumple_scope3_eval_simple = requiere_target_scope3 # Marcamos si requiere target, no si lo cumple

        brecha_reduccion = (reducción_real - umbrales['reducción_anual']) if not pd.isna(reducción_real) else np.nan

        # Cumplimiento global simplificado: ¿Se alcanzó la reducción anual requerida?
        # (La evaluación de Scope 3 es más compleja y depende de si hay target)
        cumple_total_simple = cumple_reduccion

        return {
            'Emisiones Totales 2021 (tCO₂e)': round(total_2021, 1),
            'Emisiones Totales 2022 (tCO₂e)': round(total_2022, 1),
            'Reducción Anual Observada (%)': round(reducción_real * 100, 1) if not pd.isna(reducción_real) else 'N/A',
            'Meta Reducción Anual Requerida SBTi (%)': umbrales['reducción_anual'] * 100,
            'Brecha vs Meta Reducción (%)': round(brecha_reduccion * 100, 1) if not pd.isna(brecha_reduccion) else 'N/A',
            'Cumple Meta Reducción Anual': 'Sí' if cumple_reduccion else 'No',
            '---': '---', # Separador
            'Emisiones Alcance 3 2022 (tCO₂e)': round(scope3_2022, 1),
            'Alcance 3 como % del Total 2022': round(cobertura_scope3_calc * 100, 1) if not pd.isna(cobertura_scope3_calc) else 'N/A',
            'Alcance 3 Significativo (>40%) segun SBTi': 'Sí' if scope3_es_significativo else 'No',
            'Requiere Target Específico para Alcance 3': 'Sí' if requiere_target_scope3 else 'No',
            # 'Meta Cobertura Alcance 3 SBTi (%) (si aplica)': umbrales['cobertura_scope3'] * 100 if requiere_target_scope3 else 'N/A', # Informativo
            # 'Cumple Meta Cobertura Alcance 3': 'N/A (Evaluación Simplificada)', # No podemos evaluar esto aquí
            '--- ': '---', # Separador
            'Cumplimiento Global Simplificado (Solo Reducción Anual)': 'Sí' if cumple_total_simple else 'No'
        }

    except (IndexError, ValueError, KeyError) as e:
        print(f"🚨 Error en evaluar_sbti: {e}. Verifica nombres exactos y valores numéricos en filas TOTAL/Subtotal de tus datos.")
        return {'Error': f"Fallo al calcular métricas SBTi: {e}"}

# ======================================
# Visualización Profesional (***CORREGIDA***)
# ======================================
def generar_grafico(df, ruta_guardado):
    """Gráfico de barras con contribución por categoría (CORREGIDO)."""
    required_cols = ['Categoría', 'Subcategoría', 'Emisiones 2022 Real (tCO₂e)']
    if not all(col in df.columns for col in required_cols):
        print(f"🚨 Error en generar_grafico: Faltan columnas requeridas ({required_cols}). No se generará el gráfico.")
        return

    plt.style.use('seaborn-v0_8-whitegrid') # Usar un estilo predefinido
    plt.figure(figsize=(12, 7)) # Ajustar tamaño

    df_plot = df.copy()
    df_plot['Emisiones 2022 Real (tCO₂e)'] = pd.to_numeric(df_plot['Emisiones 2022 Real (tCO₂e)'], errors='coerce')

    # Filtrar solo Alcances 1, 2, 3 y excluir totales/subtotales para el gráfico de barras
    df_plot = df_plot[
        (df_plot['Categoría'].astype(str).str.contains('Alcance [123]', case=False, na=False)) &
        (~df_plot['Categoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)) &
        (~df_plot['Subcategoría'].astype(str).str.contains('Subtotal|TOTAL', na=False, case=False)) &
        (df_plot['Emisiones 2022 Real (tCO₂e)'].notna()) &
        (df_plot['Emisiones 2022 Real (tCO₂e)'] > 0) # Graficar solo emisiones positivas
    ].copy()

    if df_plot.empty:
        print("🎨 Advertencia en generar_grafico: No hay datos válidos para graficar después de filtrar.")
        plt.close()
        return

    # *** CORRECCIÓN IMPORTANTE: ELIMINAR LA DIVISIÓN INNECESARIA ***
    # Los datos ya están (o deberían estar) en tCO₂e.
    # df_plot['Emisiones 2022 Real (tCO₂e)'] = df_plot['Emisiones 2022 Real (tCO₂e)'] / 1000 # <- ESTA LÍNEA SE ELIMINA

    df_plot['Subcategoría'] = df_plot['Subcategoría'].astype(str)
    df_plot = df_plot.sort_values('Emisiones 2022 Real (tCO₂e)', ascending=False) # Ordenar barras

    # Mapeo de colores por Alcance
    color_map = {'1': '#E63946', '2': '#457B9D', '3': '#A8DADC'} # Paleta ajustada
    # Extraer el número del alcance para asignar color
    df_plot['Alcance_Num'] = df_plot['Categoría'].astype(str).str.extract(r'(\d)', expand=False).fillna('3') # Asume 3 si no encuentra número
    bar_colors = [color_map.get(alc, '#CCCCCC') for alc in df_plot['Alcance_Num']] # Color gris si no coincide

    bars = plt.bar(
        df_plot['Subcategoría'], df_plot['Emisiones 2022 Real (tCO₂e)'], color=bar_colors
    )

    plt.title('Contribución Estimada a Emisiones 2022 por Subcategoría (tCO₂e)', fontsize=16, pad=20) # Título ajustado
    plt.xlabel('Subcategoría', fontsize=12)
    plt.ylabel('Emisiones Estimadas (tCO₂e)', fontsize=12) # Etiqueta Y ajustada
    plt.xticks(rotation=45, ha='right', fontsize=10) # Rotar menos si caben
    plt.yticks(fontsize=10)
    plt.grid(axis='y', linestyle='--', alpha=0.7) # Rejilla más visible

    # Añadir etiquetas de valor sobre las barras (formato ajustado para números grandes)
    for bar in bars:
        height = bar.get_height()
        if height > 0:
            plt.text(bar.get_x() + bar.get_width()/2., height,
                     f'{height:,.0f}', # Formato con separador de miles, sin decimales
                     ha='center', va='bottom', fontsize=9, rotation=0, fontweight='bold')

    # Crear leyenda manualmente
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color_map['1'], label='Alcance 1'),
                       Patch(facecolor=color_map['2'], label='Alcance 2'),
                       Patch(facecolor=color_map['3'], label='Alcance 3')]
    plt.legend(handles=legend_elements, title="Alcance", fontsize=10, title_fontsize=11, loc='upper right')

    plt.tight_layout() # Ajustar layout para evitar solapamientos
    try:
        plt.savefig(ruta_guardado, dpi=300, bbox_inches='tight')
        print(f"✅ Gráfico (con datos simulados) guardado en: {ruta_guardado}")
    except Exception as e:
        print(f"🚨 Error al guardar el gráfico: {e}")
    finally:
        plt.close() # Cerrar la figura para liberar memoria

# ======================================
# Generación de Reporte en Texto (Adaptado para nuevos datos/métricas)
# ======================================
def generar_reporte_texto(dict_sbti, dict_simulaciones, df_predicciones):
    """Genera un resumen ejecutivo en texto formateado."""
    reporte = "="*60 + "\n"
    reporte += "      RESUMEN EJECUTIVO ANÁLISIS DE EMISIONES (DATOS SIMULADOS)\n"
    reporte += "="*60 + "\n\n"

    reporte += "--- Evaluación SBTi Simplificada (Basada en Datos Simulados 2021-2022) ---\n"
    if dict_sbti and 'Error' not in dict_sbti:
        reporte += f"- Emisiones Totales 2021 (tCO₂e):        {dict_sbti.get('Emisiones Totales 2021 (tCO₂e)', 'N/A'):,.1f}\n"
        reporte += f"- Emisiones Totales 2022 (tCO₂e):        {dict_sbti.get('Emisiones Totales 2022 (tCO₂e)', 'N/A'):,.1f}\n"
        reporte += f"- Reducción Anual Observada (21-22):    {dict_sbti.get('Reducción Anual Observada (%)', 'N/A')} %\n"
        reporte += f"- Meta Reducción Anual Requerida SBTi:  {dict_sbti.get('Meta Reducción Anual Requerida SBTi (%)', 'N/A')} %\n"
        reporte += f"- Cumple Meta Reducción Anual:         {dict_sbti.get('Cumple Meta Reducción Anual', 'N/A')}\n\n"
        reporte += f"- Emisiones Alcance 3 2022 (tCO₂e):      {dict_sbti.get('Emisiones Alcance 3 2022 (tCO₂e)', 'N/A'):,.1f}\n"
        reporte += f"- Alcance 3 como % del Total 2022:     {dict_sbti.get('Alcance 3 como % del Total 2022', 'N/A')} %\n"
        reporte += f"- Alcance 3 Significativo (>40%):      {dict_sbti.get('Alcance 3 Significativo (>40%) segun SBTi', 'N/A')}\n"
        reporte += f"- Requiere Target Específico Alcance 3:{dict_sbti.get('Requiere Target Específico para Alcance 3', 'N/A')}\n\n"
        # reporte += f"- CUMPLIMIENTO GLOBAL (Simplificado):    {dict_sbti.get('Cumplimiento Global Simplificado (Solo Reducción Anual)', 'N/A')}\n"
    elif dict_sbti and 'Error' in dict_sbti:
         reporte += f"ERROR en la evaluación: {dict_sbti['Error']}\n"
    else:
        reporte += "Evaluación SBTi no disponible.\n"
    reporte += "-"*60 + "\n\n"

    reporte += "--- Simulación de Escenarios (Impacto en tCO₂e Totales 2022 - Base Simulable) ---\n"
    if dict_simulaciones:
        # Obtener la base de simulación para calcular reducción % (requiere pasarla o recalcularla aquí)
        # Nota: La base se imprime en la función simular_escenarios, podríamos pasarla si es necesario.
        for esc, valor in dict_simulaciones.items():
             if isinstance(valor, (int, float)):
                 reporte += f"- {esc}: {valor:,.1f} tCO₂e (Emisiones Totales Estimadas post-escenario)\n"
                 # Calcular reducción vs base simulable si tuviéramos la base aquí
             else:
                  reporte += f"- {esc}: {valor}\n" # Para mensajes de error como 'Target no encontrado'
    else:
        reporte += "Simulaciones no disponibles o no ejecutadas.\n"
    reporte += "-"*60 + "\n\n"

    reporte += "--- Predicción Simplificada 2023 (Basada en tendencia 21-22 y ajuste) ---\n"
    if df_predicciones is not None and not df_predicciones.empty:
        # Mostrar predicción por subcategoría
        reporte += "  Predicción por Subcategoría (tCO₂e):\n"
        for index, row in df_predicciones.iterrows():
             reporte += f"  - {row['Subcategoría']}: {row['Predicción 2023 (tCO₂e)']:,.1f}\n"

        total_predicho = df_predicciones['Predicción 2023 (tCO₂e)'].sum()
        reporte += f"\n- TOTAL Predicho 2023 (suma subcat.): {total_predicho:,.1f} tCO₂e\n"
    else:
        reporte += "Predicciones no disponibles.\n"
    reporte += "="*60 + "\n"

    return reporte

# ======================================
# Ejecución Principal (Adaptada para usar datos simulados)
# ======================================
if __name__ == "__main__":
    print("⏳ Iniciando análisis con DATOS SIMULADOS...")
    print("-" * 30)

    # *** CAMBIO: Usar la función para crear datos simulados ***
    # datos_originales = cargar_datos() # Comentado o eliminado
    datos_simulados = crear_datos_simulados()
    if datos_simulados is None or datos_simulados.empty:
        print("🚨 Error crítico: No se pudieron generar los datos simulados.")
        sys.exit(1)
    print(f"✅ Datos simulados generados. {len(datos_simulados)} filas creadas.")
    # print(datos_simulados.to_string()) # Descomentar para ver los datos simulados
    print("-" * 30)

    # --- El resto del flujo usa los datos simulados ---
    df_predicciones = predecir_emisiones(datos_simulados.copy()) # Pasar copia
    if df_predicciones is not None and not df_predicciones.empty:
        print(f"✅ Predicción 2023 generada para {len(df_predicciones)} subcategorías.")
    else:
        print("⚠️ No se pudo generar la predicción.")
    print("-" * 30)

    dict_simulaciones = simular_escenarios(datos_simulados.copy()) # Pasar copia
    if dict_simulaciones:
        print(f"✅ {len(dict_simulaciones)} escenarios simulados.")
    else:
        print("⚠️ No se pudieron ejecutar las simulaciones o no se encontraron targets.")
    print("-" * 30)

    dict_sbti = evaluar_sbti(datos_simulados.copy()) # Pasar copia
    if dict_sbti and 'Error' not in dict_sbti:
        print("✅ Evaluación SBTi (simplificada, datos simulados) completada.")
    elif dict_sbti and 'Error' in dict_sbti:
        print(f"⚠️ Error en la evaluación SBTi: {dict_sbti['Error']}")
    else:
        print("⚠️ No se pudo completar la evaluación SBTi.")
    print("-" * 30)

    print(f"💾 Guardando resultados tabulares (datos simulados) en: {ARCHIVO_RESULTADOS_EXCEL}")
    try:
        with pd.ExcelWriter(ARCHIVO_RESULTADOS_EXCEL, engine='openpyxl') as writer:
            # Guardar los datos simulados originales usados
            datos_simulados.to_excel(writer, sheet_name='Datos_Simulados_Input', index=False)
            print("   - Hoja 'Datos_Simulados_Input' guardada.")

            if df_predicciones is not None and not df_predicciones.empty:
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

        print(f"✅ Resultados tabulares (datos simulados) guardados exitosamente.")

    except Exception as e:
        print(f"🚨 Error Crítico al guardar resultados en Excel: {str(e)}")
    print("-" * 30)

    print("🎨 Generando gráfico de contribución (con datos simulados)...")
    # *** CAMBIO: Pasar los datos simulados a la función de gráfico ***
    generar_grafico(datos_simulados, ARCHIVO_GRAFICO)
    print("-" * 30)

    print("📄 Generando reporte resumen en consola (con datos simulados)...")
    reporte_final = generar_reporte_texto(dict_sbti, dict_simulaciones, df_predicciones)
    print("\n" + reporte_final)
    print("-" * 30)
    print("🏁 Análisis completado (con datos simulados).")
