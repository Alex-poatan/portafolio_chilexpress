import pandas as pd
import os

# --- Configuración básica ---
RUTA_EXCEL = os.path.join("data", "raw", "Planilla de datos de Huella de Carbono Chilexpress 2022.xlsx")

# --- Leer todas las hojas ---
try:
    # Cargar archivo Excel
    excel_file = pd.ExcelFile(RUTA_EXCEL)
    hojas = excel_file.sheet_names  # Lista de 24 hojas

    # Almacenar DataFrames en un diccionario
    datos_hojas = {hoja: excel_file.parse(hoja) for hoja in hojas}

    # Mostrar resumen
    print("✅ ¡Hojas leídas exitosamente!")
    print(f"📂 Hojas procesadas: {len(datos_hojas)}")
    print("🔍 Nombres de las hojas:", list(datos_hojas.keys()))

except Exception as e:
    print(f"❌ Error al leer el archivo: {str(e)}")
finally:
    # Liberar recursos
    if 'excel_file' in locals():
        del excel_file
