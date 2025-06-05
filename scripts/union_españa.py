import os
import pandas as pd

# Rutas
input_dir = r'C:\Users\SERVIDOR3015\Desktop\unificacion-espa-a-python\data\unionEspaña\input'
output_dir = r'C:\Users\SERVIDOR3015\Desktop\unificacion-espa-a-python\data\unionEspaña\output'
csv_output_path = os.path.join(output_dir, 'union_españa12.csv')

# Crear lista para guardar los dataframes
all_dataframes = []

# Leer todos los archivos Excel del directorio
for file in os.listdir(input_dir):
    if file.endswith('.xlsx') or file.endswith('.xls'):
        file_path = os.path.join(input_dir, file)
        try:
            df = pd.read_excel(file_path)
            all_dataframes.append(df)
            print(f'Archivo cargado: {file}')
        except Exception as e:
            print(f'Error al leer {file}: {e}')

# Unir todos los DataFrames
if all_dataframes:
    full_df = pd.concat(all_dataframes, ignore_index=True)

    # Eliminar duplicados por 'cid' y 'nombre'
    if 'cid' in full_df.columns and 'nombre' in full_df.columns:
        full_df.drop_duplicates(subset=['cid', 'nombre'], inplace=True)
        print("Duplicados eliminados por 'cid' y 'nombre'.")

    # Asegurarse de que existan las columnas necesarias
    required_cols = ['categoría', 'rating', 'nº reviews', 'reviewstext', 'email', 'website']
    for col in required_cols:
        if col not in full_df.columns:
            full_df[col] = None  # Añadir columna si no existe

    def is_filled(val):
        return pd.notna(val) and str(val).strip() != ""

    def keep_row(row):
        has_main_data = any(is_filled(row[col]) for col in ['categoría', 'rating', 'nº reviews', 'reviewstext'])
        has_contact_info = is_filled(row['email']) or is_filled(row['website'])
        return has_main_data and has_contact_info

    # Filtrar el DataFrame
    original_count = len(full_df)
    full_df = full_df[full_df.apply(keep_row, axis=1)]
    filtered_count = len(full_df)
    removed_count = original_count - filtered_count
    print(f"Filtrado aplicado. Filas eliminadas: {removed_count}")

    # Guardar en CSV
    full_df.to_csv(csv_output_path, index=False, encoding='utf-8-sig')
    print(f'Archivo CSV guardado: {csv_output_path}')

    # Dividir en archivos Excel si hay más de 500,000 filas
    rows_per_file = 900_000
    num_files = (len(full_df) // rows_per_file) + 1

    for i in range(num_files):
        start = i * rows_per_file
        end = start + rows_per_file
        partial_df = full_df.iloc[start:end]
        excel_output_path = os.path.join(output_dir, f'union_españa_excel{i+1}.xlsx')
        partial_df.to_excel(excel_output_path, index=False)
        print(f'Archivo Excel guardado: {excel_output_path}')
else:
    print("No se encontraron archivos Excel para procesar.")
