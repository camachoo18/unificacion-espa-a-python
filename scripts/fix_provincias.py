import pandas as pd
import os

# Ruta al archivo original
archivo = r"C:\Users\SERVIDOR3015\Desktop\unificacion-espa-a-python\provincias_limpias\input\Fuenlabrada_limpio.xlsx"

# Cargar el Excel (asegúrate de que la hoja se llame 'data')
df = pd.read_excel(archivo, sheet_name='Sheet1')

# Filtrar filas donde la columna 'localidad' sea exactamente 'Fuenlabrada' (ignorando mayúsculas y espacios)
df_filtrado = df[df['localidad'].str.strip().str.lower() == 'fuenlabrada']

# Ruta de guardado del nuevo archivo
salida = r"C:\Users\SERVIDOR3015\Desktop\unificacion-espa-a-python\provincias_limpias\input\filtrado_fuenlabrada.xlsx"

# Guardar el archivo filtrado
df_filtrado.to_excel(salida, index=False)

print("Filtrado completado. Archivo guardado en:", salida)
