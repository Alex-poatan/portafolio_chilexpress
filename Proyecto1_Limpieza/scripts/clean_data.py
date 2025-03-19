"""
Script para limpiar y procesar datos de huella de carbono de Chilexpress.
Autor: Tu Nombre
Fecha: Agosto 2023
"""
import os
import pandas as pd

# --- Configuración ---
NOMBRE_ARCHIVO = "Planilla de datos de Huella de Carbono Chilexpress 2022.xlsx"

# --- Ruta absoluta para diagnóstico ---
ruta_absoluta = os.path.join(
    os.environ["USERPROFILE"],
    "OneDrive",
    "Escritorio",
    "portafolio_chilexpress",
    "Proyecto1_Limpieza",
    "data",
    "raw",
    NOMBRE_ARCHIVO
)

print(f"\n[DEBUG] Buscando archivo en:\n{ruta_absoluta}\n")

# --- Verificar si el archivo existe ---
if not os.path.exists(ruta_absoluta):
    print("❌ Error: El archivo NO existe en la ruta indicada.")
    print("Posibles soluciones:")
    print("- Verifica que el archivo esté en la carpeta 'Proyecto1_Limpieza/data/raw/'.")
    print("- Asegúrate de que el nombre coincida (incluyendo espacios y mayúsculas).")
else:
    print("✅ Archivo encontrado! Leyendo datos...\n")

    # Leer el archivo
    try:
        df = pd.read_excel(ruta_absoluta, sheet_name='2022 Estimado')
        print("✅ Datos leídos correctamente!")
        print("\nPrimeras filas:")
        print(df.head())
    except Exception as e:  # pylint: disable=W0718
        print("❌ Error al leer el archivo:", str(e))
