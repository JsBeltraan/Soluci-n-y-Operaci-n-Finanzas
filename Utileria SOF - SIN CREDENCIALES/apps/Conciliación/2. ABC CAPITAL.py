import pandas as pd
import os
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
import logging
import sys

# Configuración del log (bitácora)
logging.basicConfig(filename='procesamiento_log.txt', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Crear una aplicación de PyQt5
app = QApplication(sys.argv)

# Mostrar una alerta inicial para informar al usuario
QMessageBox.information(
    None,
    "Instrucciones",
    "Seleccione una carpeta que contenga archivos XLSX u ODS de ABC Capital para procesar."
)

# Abrir un diálogo para seleccionar la carpeta
folder_path = QFileDialog.getExistingDirectory(None, "Seleccionar Carpeta con Archivos Excel")

# Validar que la carpeta contenga al menos un archivo .xlsx o .ods
def validar_carpeta(path):
    """
    Función para validar que la carpeta seleccionada contenga archivos válidos (.xlsx o .ods).
    """
    if not path:  # Si no se seleccionó ninguna carpeta
        return False
    archivos_validos = [f for f in os.listdir(path) if f.endswith(('.xlsx', '.ods'))]
    return len(archivos_validos) > 0

# Si la carpeta no contiene archivos válidos, mostrar una alerta
if not validar_carpeta(folder_path):
    QMessageBox.warning(
        None,
        "Carpeta Inválida",
        "Debe seleccionar una carpeta que contenga archivos XLSX u ODS de ABC Capital."
    )
    logging.warning("No se seleccionó una carpeta válida con archivos XLSX u ODS.")
    sys.exit()

# Inicializar una lista vacía para almacenar los DataFrames
dataframes = []

# Iterar sobre cada archivo en la carpeta
for filename in os.listdir(folder_path):
    if filename.endswith(('.xlsx', '.ods')):  # Filtrar solo los archivos con extensiones válidas
        file_path = os.path.join(folder_path, filename)
        try:
            # Leer el archivo completo
            temp_df = pd.read_excel(file_path, sheet_name='Reporte General', dtype=str)

            # Buscar la fila donde aparece la columna "EMPRESA"
            empresa_row = temp_df[temp_df.apply(lambda row: row.str.contains('EMPRESA', na=False).any(), axis=1)].index

            if empresa_row.empty:
                # Si no se encuentra la columna "EMPRESA", registrar un error y omitir el archivo
                logging.error(f"Columna 'EMPRESA' no encontrada en {filename}. Se omitirá este archivo.")
                continue

            # Determinar la fila de inicio de los datos reales
            start_row = empresa_row[0] + 1

            # Leer desde la fila donde comienzan los datos
            df = pd.read_excel(file_path, sheet_name='Reporte General', skiprows=start_row, dtype=str).replace("'", "")
            df.columns = df.columns.str.strip()  # Limpiar los nombres de las columnas

            # Validar columnas requeridas
            required_columns = ['EMPRESA', 'FECHA', 'MOVIMIENTOS', 'IMPORTE', 'BANCO']
            if not all(col in df.columns for col in required_columns):
                logging.error(f"Columnas requeridas no encontradas en {filename}. Se omitirá este archivo.")
                continue

            # Procesar datos
            df = df[df['BANCO'].notna() & (df['BANCO'] != '')]  # Filtrar filas donde 'BANCO' no sea nulo o vacío
            df = df[df['FECHA'].str.match(r'\d{4}-\d{2}-\d{2}', na=False)]  # Validar formato de fecha
            df['Fecha'] = pd.to_datetime(df['FECHA'], format='%Y-%m-%d', errors='coerce')  # Convertir a datetime
            df = df.dropna(subset=['Fecha'])  # Eliminar filas con fechas inválidas
            df['AÑO'] = df['Fecha'].dt.year  # Extraer el año de la fecha
            df['MES'] = df['Fecha'].dt.month  # Extraer el mes de la fecha
            df['SERVICIO'] = 'ABC CAPITAL'  # Agregar una columna con el servicio
            df['MOVIMIENTOS'] = pd.to_numeric(df['MOVIMIENTOS'], errors='coerce')  # Convertir movimientos a numérico

            # Limpiar y convertir 'IMPORTE'
            df['IMPORTE'] = df['IMPORTE'].astype(str).replace({'\$': '', ',': ''}, regex=True)  # Quitar caracteres no numéricos
            df['IMPORTE'] = pd.to_numeric(df['IMPORTE'], errors='coerce')  # Convertir a numérico

            # Seleccionar columnas relevantes
            df = df[['AÑO', 'MES', 'SERVICIO', 'MOVIMIENTOS', 'IMPORTE']]

            # Agregar DataFrame a la lista
            dataframes.append(df)
        except Exception as e:
            # Registrar cualquier error ocurrido durante el procesamiento del archivo
            logging.error(f"Error al procesar el archivo {filename}: {e}")
            continue

# Combinar y guardar resultados
if dataframes:
    combined_df = pd.concat(dataframes, ignore_index=True)  # Combinar todos los DataFrames

    # Agrupar por AÑO, MES y SERVICIO, sumar movimientos e importes
    summary_df = combined_df.groupby(['AÑO', 'MES', 'SERVICIO'], as_index=False).agg({
        'MOVIMIENTOS': 'sum',
        'IMPORTE': 'sum'
    })

    # Formatear 'IMPORTE'
    summary_df['IMPORTE'] = summary_df['IMPORTE'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")

    # Guardar en un archivo de Excel
    output_file = os.path.join(folder_path, 'REPORTE_CONSOLIDADO_ABC_CAPITAL.xlsx')
    summary_df.to_excel(output_file, index=False)

    logging.info(f"Proceso completado. Archivo guardado en: {output_file}")

    # Mostrar mensaje de éxito
    QMessageBox.information(
        None,
        "Proceso Completado",
        f"El proceso ha finalizado exitosamente. El archivo consolidado se guardó en:\n{output_file}"
    )
else:
    # Si no se procesaron archivos, registrar un mensaje y mostrar alerta
    logging.warning("No se procesaron archivos debido a errores.")
    QMessageBox.warning(
        None,
        "Sin Resultados",
        "No se procesaron archivos debido a errores o porque no se encontraron datos válidos."
    )

# Finalizar la aplicación
sys.exit(app.exec_())
