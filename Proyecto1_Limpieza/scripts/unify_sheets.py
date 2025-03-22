"Unificación de datos"

import os
import pandas as pd


# --- Configuración ---
RUTA_BASE = os.path.dirname(__file__)
RUTA_EXCEL = os.path.join(RUTA_BASE, "..", "data", "raw", "Planilla de datos de Huella de Carbono Chilexpress 2022.xlsx")
COLUMNAS_ESPERADAS = ["categoria", "fecha", "toneladas_co2", "ruta"]  # Ajusta según tus datos

def encontrar_fila_inicio(hoja, palabra_clave="toneladas_co2"):
    """Busca la fila donde están los encabezados reales."""
    try:
        df_prueba = pd.read_excel(RUTA_EXCEL, sheet_name=hoja, header=None, nrows=10)
        for idx, fila in df_prueba.iterrows():
            if any(str(celda).lower() == palabra_clave for celda in fila.values):
                print(f"📌 Encabezados encontrados en fila {idx} de la hoja '{hoja}'")
                return idx
        print(f"⚠️ No se encontró '{palabra_clave}' en la hoja '{hoja}'. Usando header=0.")
        return 0
    except Exception as e:
        print(f"🔥 Error al buscar encabezados en '{hoja}': {str(e)}")
        return 0

def unificar_dataframe():
    try:
        print("\n🔍 Iniciando proceso...")

        # Verificar si el archivo existe
        if not os.path.exists(RUTA_EXCEL):
            raise FileNotFoundError(f"Archivo no encontrado: {RUTA_EXCEL}")

        # Cargar Excel
        excel_file = pd.ExcelFile(RUTA_EXCEL)
        hojas = excel_file.sheet_names
        print(f"📂 Hojas detectadas: {hojas}")

        dataframes = []
        for hoja in hojas:
            try:
                print(f"\n📄 Procesando hoja: '{hoja}'...")

                # 1. Encontrar fila de inicio
                fila_inicio = encontrar_fila_inicio(hoja)

                # 2. Leer datos
                df = pd.read_excel(excel_file, sheet_name=hoja, header=fila_inicio)
                print(f"   ✅ Datos leídos ({len(df)} filas)")

                # 3. Limpiar columnas
                df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
                df.columns = df.columns.astype(str)
                df = df.loc[:, ~df.columns.str.contains("Unnamed|Notas|Comentario", case=False, na=False)]
                print("   ✅ Columnas limpiadas")

                # 4. Estandarizar nombres
                mapeo_columnas = {
                    "toneladas co2": "toneladas_co2",
                    "emisiones co2": "toneladas_co2",
                    "fecha": "fecha",
                    "ruta": "ruta"
                }
                df = df.rename(columns=lambda x: mapeo_columnas.get(x.strip().lower(), x))
                df["categoria"] = hoja.strip().lower().replace(" ", "_")
                print("   ✅ Nombres estandarizados")

                # 5. Filtrar columnas
                columnas_validas = [col for col in COLUMNAS_ESPERADAS if col in df.columns]
                if not columnas_validas:
                    print(f"⚠️ Hoja '{hoja}' no tiene columnas válidas. Saltando...")
                    continue
                df = df[columnas_validas]
                dataframes.append(df)
                print(f"   ✅ Hoja añadida al DataFrame final")

            except Exception as e:
                print(f"❌ Error procesando '{hoja}': {str(e)}")
                continue

        # Unificar DataFrames
        if not dataframes:
            raise ValueError("🚨 Ninguna hoja válida fue procesada.")

        df_unificado = pd.concat(dataframes, ignore_index=True)
        print("\n🎉 ¡Proceso completado con éxito!")
        print(f"- Filas totales: {len(df_unificado)}")
        print(f"- Columnas finales: {df_unificado.columns.tolist()}")

        # Exportar a CSV
        RUTA_SALIDA = os.path.join(RUTA_BASE, "..", "data", "cleaned", "datos_finales.csv")
        df_unificado.to_csv(RUTA_SALIDA, index=False, encoding="utf-8-sig")
        print(f"\n💾 Datos guardados en: {RUTA_SALIDA}")

        return df_unificado

    except Exception as e:
        print(f"\n❌ Error crítico: {str(e)}")
        return None

if __name__ == "__main__":
    df_final = unificar_dataframe()
    if df_final is not None:
        print("\n✅ ¡Ejecución exitosa! Verifica el archivo CSV generado.")
    else:
        print("\n⚠️ La ejecución terminó con errores. Revisa los mensajes anteriores.")
