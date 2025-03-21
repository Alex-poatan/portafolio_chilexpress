import pandas as pd
import os

# --- Configuración corregida ---
RUTA_BASE = os.path.dirname(__file__)  # Ruta del script actual
RUTA_EXCEL = os.path.join(RUTA_BASE, "..", "data", "raw", "Planilla de datos de Huella de Carbono Chilexpress 2022.xlsx")  # Ajustado

# --- Leer todas las hojas ---
try:
    excel_file = pd.ExcelFile(RUTA_EXCEL)
    hojas = excel_file.sheet_names

    datos_hojas = {hoja: excel_file.parse(hoja) for hoja in hojas}

    print("✅ ¡Hojas leídas exitosamente!")
    print("📂 Hojas procesadas:", len(datos_hojas))
    print("🔍 Nombres de las hojas:", list(datos_hojas.keys()))

except Exception as e:
    print(f"❌ Error al leer el archivo: {str(e)}")
    print(f"⚠️ Ruta intentada: {RUTA_EXCEL}")  # Debug: muestra la ruta
finally:
    if 'excel_file' in locals():
        del excel_file
