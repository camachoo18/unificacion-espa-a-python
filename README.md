# Documentación de Scripts - unificacion-espa-a-python/scripts/

Este documento describe la funcionalidad y propósito de cada script encontrado en la carpeta `/scripts/`.

---

## 1. `fix_provincias.py`

**Propósito:**  
Filtra un archivo Excel para quedarse únicamente con las filas de la localidad "Fuenlabrada" y guarda el resultado en un nuevo archivo Excel.

**¿Qué hace?**
- Carga un archivo Excel específico.
- Filtra las filas donde la columna `localidad` es exactamente "Fuenlabrada" (ignorando mayúsculas y espacios).
- Guarda el resultado filtrado en otro archivo Excel.
- Muestra un mensaje indicando la ubicación del archivo generado.

**Uso típico:**  
Extraer únicamente los datos relativos a "Fuenlabrada" de un conjunto de datos más amplio de localidades.

---

## 2. `limpieza_provincias.py`

**Propósito:**  
Realiza una limpieza avanzada de archivos Excel de provincias, eliminando columnas y filas innecesarias, deduplicando y añadiendo una puntuación personalizada para cada fila.

**¿Qué hace?**
- Recorre todos los archivos Excel de un directorio de entrada.
- Elimina columnas irrelevantes o redundantes.
- Elimina filas duplicadas en base a las columnas `CID` y `nombre`.
- Elimina filas vacías o que no contengan información relevante.
- Añade una columna `ratingCompany` que puntúa la completitud de cada registro evaluando si tiene email, website, rating y teléfono.
- Guarda los archivos ya limpios en un directorio de salida.
- Informa por consola sobre el progreso y las acciones realizadas para cada archivo.

**Uso típico:**  
Preparar y limpiar bases de datos provinciales para un posterior proceso de unificación, asegurando calidad y homogeneidad en los datos.

---

## 3. `union_españa.py`

**Propósito:**  
Une todos los archivos Excel de provincias limpias en un único DataFrame, aplica filtros adicionales y guarda el resultado como CSV y, si es necesario, como varios archivos Excel particionados.

**¿Qué hace?**
- Lee todos los archivos Excel de un directorio de entrada y los concatena.
- Elimina duplicados basados en las columnas `cid` y `nombre`.
- Se asegura de que existan ciertas columnas necesarias (las crea si faltan).
- Filtra filas que no tengan información relevante en al menos una columna principal y sin datos de contacto (email/website).
- Guarda el DataFrame unido como un archivo CSV.
- Si el archivo supera cierto tamaño, lo divide en varios archivos Excel de hasta 900,000 filas cada uno.
- Informa por consola el avance y los archivos generados.

**Uso típico:**  
Unificar y consolidar todas las provincias en una sola base de datos nacional depurada, lista para análisis o cargas masivas.

---

> **Nota:** Todos los scripts están diseñados para procesar archivos Excel y requieren que la estructura de las carpetas y los nombres de las columnas sean coherentes con la configuración interna de cada script.
