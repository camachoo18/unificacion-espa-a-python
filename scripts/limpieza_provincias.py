import os
import pandas as pd
import numpy as np  # Para pd.NA y np.nan

# --- Configuración de Directorios ---
INPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'unionEspaña', 'input')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'unionEspaña', 'output')

# --- Configuración de Columnas (AJUSTAR SEGÚN LOS NOMBRES EN TUS ARCHIVOS EXCEL) ---
# Estos nombres se buscarán de forma insensible a mayúsculas/minúsculas en los archivos Excel.
COLUMNAS_DEDUPLICACION = ['CID', 'nombre']
COLUMNA_ORIGINAL_KW = 'ORIGINAL KW'
COLUMNAS_A_ELIMINAR_ADICIONALES = ["descripción","Descripción", "id", " Day N.", "Saved", "hash", "title", "EmailVerificado"]
COLUMNAS_PARA_FILA_VACIA_CHECK = ["Rating", " Nº Reviews", "ReviewsText", "Horario", "Provincia"]
COLUMNA_EMAIL = 'email'
COLUMNA_WEBSITE = 'Website'
COLUMNA_ORIGINAL_RATING_PARA_PUNTUACION = 'Rating'
COLUMNA_TELEFONO = 'Telefono'


# --- Fin de Configuración de Columnas ---

def get_actual_column_name(configured_name, df_column_list):
    """
    Encuentra el nombre de columna real (sensible a mayúsculas/minúsculas) en la lista de columnas de un DataFrame
    que coincida con configured_name de forma insensible a mayúsculas/minúsculas.
    Devuelve el nombre de columna real o None si no se encuentra.
    """
    if configured_name is None:
        return None
    configured_name_lower = configured_name.lower()
    for actual_col in df_column_list:
        if actual_col.lower() == configured_name_lower:
            return actual_col
    return None


def es_vacio(valor):
    """Verifica si un valor es considerado vacío (NaN, None, o string vacío/solo espacios)."""
    if pd.isna(valor):
        return True
    if isinstance(valor, str) and not valor.strip():
        return True
    return False


def calcular_rating_company(row, actual_cols_info):
    """Calcula la puntuación para la columna 'ratingCompany' usando nombres de columna reales."""
    score = 10

    # Penalización por Email
    actual_email_col = actual_cols_info['email']
    if actual_email_col and es_vacio(row[actual_email_col]):
        score -= 2
    elif not actual_email_col:  # Si la columna (insensible a mayúsculas) no se encontró en el DF
        score -= 2

    # Penalización por Website
    actual_website_col = actual_cols_info['website']
    if actual_website_col and es_vacio(row[actual_website_col]):
        score -= 2
    elif not actual_website_col:
        score -= 2

    # Penalización por Rating original
    actual_rating_col = actual_cols_info['original_rating']
    if actual_rating_col and es_vacio(row[actual_rating_col]):
        score -= 1
    elif not actual_rating_col:
        score -= 1

    # Penalización por Teléfono
    actual_telefono_col = actual_cols_info['telefono']
    if actual_telefono_col and es_vacio(row[actual_telefono_col]):
        score -= 1
    elif not actual_telefono_col:
        score -= 1

    return max(0, score)  # Asegura que la puntuación no sea menor que 0


def limpiar_dataframe_provincia(df, nombre_archivo_original):
    """
    Limpia un DataFrame de una provincia específica según los requisitos.
    """
    print(f"\nProcesando archivo: {nombre_archivo_original}")
    print(f"Filas iniciales: {len(df)}")
    df_limpio = df.copy()
    df_column_list_original = df_limpio.columns.tolist()  # Lista de columnas originales del DF actual

    # 0. Obtener nombres reales de columnas configuradas para eliminación
    actual_col_original_kw = get_actual_column_name(COLUMNA_ORIGINAL_KW, df_column_list_original)
    actual_cols_a_eliminar_adicionales = [
        get_actual_column_name(col, df_column_list_original) for col in COLUMNAS_A_ELIMINAR_ADICIONALES
    ]
    actual_cols_a_eliminar_adicionales = [col for col in actual_cols_a_eliminar_adicionales if col]  # Filtrar Nones

    todas_actual_columnas_a_eliminar = actual_cols_a_eliminar_adicionales
    if actual_col_original_kw:
        todas_actual_columnas_a_eliminar.append(actual_col_original_kw)

    # Mantener un conjunto único en caso de solapamiento (poco probable con esta config)
    todas_actual_columnas_a_eliminar = list(set(todas_actual_columnas_a_eliminar))

    # 1. Eliminar columnas especificadas (usando nombres reales)
    columnas_fisicamente_eliminadas = []
    if todas_actual_columnas_a_eliminar:
        cols_a_dropear = [col for col in todas_actual_columnas_a_eliminar if col in df_limpio.columns]
        if cols_a_dropear:
            df_limpio.drop(columns=cols_a_dropear, inplace=True)
            columnas_fisicamente_eliminadas = cols_a_dropear
            print(f"Columnas eliminadas: {columnas_fisicamente_eliminadas}")

    if not columnas_fisicamente_eliminadas:
        # Informar si ninguna de las configuradas (insensible a mayúsculas) se encontró
        nombres_config_originales = COLUMNAS_A_ELIMINAR_ADICIONALES + [COLUMNA_ORIGINAL_KW]
        print(
            f"Ninguna de las columnas configuradas para eliminación ({', '.join(nombres_config_originales)}) fue encontrada (búsqueda insensible a mayúsculas/minúsculas).")

    # Actualizar lista de columnas después de eliminaciones
    df_column_list_actual = df_limpio.columns.tolist()

    # 2. Eliminar duplicados basados en 'cid' y 'nombre' (usando nombres reales)
    actual_cols_deduplicacion = [
        get_actual_column_name(col, df_column_list_actual) for col in COLUMNAS_DEDUPLICACION
    ]
    actual_cols_deduplicacion_existentes = [col for col in actual_cols_deduplicacion if col]

    if len(actual_cols_deduplicacion_existentes) == len(COLUMNAS_DEDUPLICACION):
        df_limpio.drop_duplicates(subset=actual_cols_deduplicacion_existentes, keep='first', inplace=True)
        print(f"Filas después de eliminar duplicados por {actual_cols_deduplicacion_existentes}: {len(df_limpio)}")
    elif actual_cols_deduplicacion_existentes:  # Si se encontró al menos una de las de deduplicación
        print(
            f"Advertencia: Solo se encontraron algunas columnas para deduplicación (insensible a mayúsculas): {actual_cols_deduplicacion_existentes}. Se usarán estas.")
        df_limpio.drop_duplicates(subset=actual_cols_deduplicacion_existentes, keep='first', inplace=True)
        print(f"Filas después de eliminar duplicados por {actual_cols_deduplicacion_existentes}: {len(df_limpio)}")
    else:
        print(
            f"Advertencia: Ninguna de las columnas para deduplicación ({', '.join(COLUMNAS_DEDUPLICACION)}) fue encontrada (búsqueda insensible a mayúsculas). Saltando este paso.")

    # 3. Lógica de eliminación de filas (actualizada, usando nombres reales)
    filas_antes_eliminacion_condicional = len(df_limpio)
    actual_cols_para_fila_vacia_check = [
        get_actual_column_name(col, df_column_list_actual) for col in COLUMNAS_PARA_FILA_VACIA_CHECK
    ]
    actual_cols_para_fila_vacia_check_existentes = [col for col in actual_cols_para_fila_vacia_check if col]

    actual_email_col_para_check = get_actual_column_name(COLUMNA_EMAIL, df_column_list_actual)
    actual_website_col_para_check = get_actual_column_name(COLUMNA_WEBSITE, df_column_list_actual)

    if actual_cols_para_fila_vacia_check_existentes:
        condicion_todas_vacias = pd.Series([True] * len(df_limpio), index=df_limpio.index)
        for col_check_actual in actual_cols_para_fila_vacia_check_existentes:
            condicion_todas_vacias &= df_limpio[col_check_actual].apply(es_vacio)

        filas_a_eliminar_final = condicion_todas_vacias
        if actual_email_col_para_check and actual_website_col_para_check:
            condicion_tiene_email_y_website = (~df_limpio[actual_email_col_para_check].apply(es_vacio)) & \
                                              (~df_limpio[actual_website_col_para_check].apply(es_vacio))
            filas_a_eliminar_final &= ~condicion_tiene_email_y_website
        elif actual_email_col_para_check:  # Solo email existe (o fue encontrado)
            condicion_tiene_email = ~df_limpio[actual_email_col_para_check].apply(es_vacio)
            filas_a_eliminar_final &= ~condicion_tiene_email
        elif actual_website_col_para_check:  # Solo website existe (o fue encontrado)
            condicion_tiene_website = ~df_limpio[actual_website_col_para_check].apply(es_vacio)
            filas_a_eliminar_final &= ~condicion_tiene_website

        df_limpio = df_limpio[~filas_a_eliminar_final]
        print(f"Filas después de eliminación condicional (Rating, etc. vacíos Y sin Email/Website): {len(df_limpio)}")
        if len(df_limpio) < filas_antes_eliminacion_condicional:
            print(f"  ({filas_antes_eliminacion_condicional - len(df_limpio)} filas eliminadas en este paso)")
    else:
        print(
            f"Advertencia: Ninguna de las columnas para el chequeo de filas vacías ({', '.join(COLUMNAS_PARA_FILA_VACIA_CHECK)}) fue encontrada (búsqueda insensible a mayúsculas). Saltando este paso.")

    # 4. Añadir columna 'ratingCompany' (usando nombres reales)
    actual_cols_info_para_rating = {
        'email': get_actual_column_name(COLUMNA_EMAIL, df_column_list_actual),
        'website': get_actual_column_name(COLUMNA_WEBSITE, df_column_list_actual),
        'original_rating': get_actual_column_name(COLUMNA_ORIGINAL_RATING_PARA_PUNTUACION, df_column_list_actual),
        'telefono': get_actual_column_name(COLUMNA_TELEFONO, df_column_list_actual)
    }

    if not df_limpio.empty:
        df_limpio['ratingCompany'] = df_limpio.apply(
            lambda row: calcular_rating_company(row, actual_cols_info_para_rating), axis=1)
        print(f"Columna 'ratingCompany' añadida.")
    else:
        print("DataFrame vacío antes de añadir 'ratingCompany'. Saltando.")

    print(f"Filas finales en '{nombre_archivo_original}': {len(df_limpio)}")
    return df_limpio


def procesar_archivos_provincia():
    """
    Itera sobre los archivos Excel en el directorio de entrada, los limpia y los guarda en el directorio de salida.
    """
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"Directorio de salida creado: {OUTPUT_DIR}")

    archivos_excel_input = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx') and not f.startswith('~$')]

    if not archivos_excel_input:
        print(f"No se encontraron archivos Excel en '{INPUT_DIR}'.")
        return

    print(f"Se encontraron {len(archivos_excel_input)} archivos Excel para procesar.")

    for nombre_archivo in archivos_excel_input:
        ruta_completa_archivo_input = os.path.join(INPUT_DIR, nombre_archivo)
        try:
            df_original = pd.read_excel(ruta_completa_archivo_input, engine='openpyxl')
            df_limpio = limpiar_dataframe_provincia(df_original, nombre_archivo)

            nombre_base_sin_ext = os.path.splitext(nombre_archivo)[0]
            partes_nombre = nombre_base_sin_ext.split('-')

            nombre_provincia_extraido = "desconocido"
            # Intenta extraer el nombre de la provincia de forma más robusta al caso
            nombre_メガ_parte = partes_nombre[0].upper()
            if nombre_メガ_parte.startswith("MEGA "):
                nombre_provincia_extraido = partes_nombre[0][5:].strip()
            elif nombre_メガ_parte.startswith("MEGA"):
                nombre_provincia_extraido = partes_nombre[0][4:].strip()
            else:
                nombre_provincia_extraido = partes_nombre[0].strip()

            if not nombre_provincia_extraido:  # Fallback
                nombre_provincia_extraido = f"datos_limpios_{nombre_base_sin_ext}"

            nombre_archivo_salida = f"{nombre_provincia_extraido}_limpio.xlsx"
            ruta_completa_archivo_salida = os.path.join(OUTPUT_DIR, nombre_archivo_salida)

            if not df_limpio.empty:
                df_limpio.to_excel(ruta_completa_archivo_salida, index=False, engine='openpyxl')
                print(f"Archivo limpio guardado en: {ruta_completa_archivo_salida}")
            else:
                print(
                    f"El DataFrame para '{nombre_archivo}' está vacío después de la limpieza. No se guardará archivo.")

        except Exception as e:
            print(f"ERROR procesando el archivo '{nombre_archivo}': {e}")
            import traceback
            print(traceback.format_exc())  # Imprime el stack trace completo


if __name__ == "__main__":
    print(f"Iniciando script de limpieza de Excel por provincia...")
    print(f"Directorio de entrada: {INPUT_DIR}")
    print(f"Directorio de salida: {OUTPUT_DIR}\n")
    procesar_archivos_provincia()
    print("\nProceso completado.")