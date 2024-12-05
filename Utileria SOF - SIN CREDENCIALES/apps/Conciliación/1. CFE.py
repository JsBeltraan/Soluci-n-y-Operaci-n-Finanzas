import pandas as pd
import os
from tkinter import Tk, messagebox
from tkinter.filedialog import askdirectory
import logging
from datetime import timedelta

# Configuración del log (bitácora)
logging.basicConfig(filename='procesamiento_log.txt', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Crear una ventana raíz de Tkinter (no visible)
root = Tk()
root.withdraw()

# Mostrar mensaje para seleccionar la carpeta
messagebox.showinfo("Alerta", "Seleccionar carpeta con archivos xlsx u ods de CFE")

# Abrir un diálogo para seleccionar la carpeta
folder_path = askdirectory(title='Seleccionar Carpeta con Archivos Excel')

# Lista para almacenar los DataFrames procesados
dataframes = []

# Función para leer archivos .ods
def read_ods(file_path):
    """
    Lee archivos .ods utilizando pandas y el motor 'odf'.
    """
    try:
        return pd.read_excel(file_path, engine='odf', dtype=str)
    except Exception as e:
        logging.error(f"Error al leer el archivo .ods {file_path}: {e}")
        return None

# Iterar sobre los archivos en la carpeta seleccionada
for filename in os.listdir(folder_path):
    if filename.endswith('.xls') or filename.endswith('.xlsx') or filename.endswith('.ods'):
        # Construir la ruta completa del archivo
        file_path = os.path.join(folder_path, filename)
        
        try:
            # Leer el archivo según su tipo
            if filename.endswith('.ods'):
                temp_df = read_ods(file_path)
            else:
                temp_df = pd.read_excel(file_path, sheet_name='Reporte CFE', dtype=str)
            
            # Buscar fila con la columna "Movimientos Coppel" en un rango específico
            search_rows = temp_df.iloc[13:17]  # Filas 14 a 17 (base cero)
            movimientos_row = search_rows[search_rows.apply(lambda row: row.str.contains('Movimientos Coppel', na=False).any(), axis=1)].index

            # Validar si no se encontró la fila
            if movimientos_row.empty:
                logging.error(f"Columna 'Movimientos Coppel' no encontrada en {filename}. Se omitirá este archivo.")
                continue

            # Definir la fila de inicio para leer datos
            start_row = movimientos_row[0] + 1
            logging.info(f"Archivo {filename} procesado desde la fila {start_row}.")
            
            # Leer el archivo a partir de la fila de datos
            df = pd.read_excel(file_path, sheet_name='Reporte CFE', skiprows=start_row, dtype=str).replace("'", "")
            
            # Limpiar nombres de columnas
            df.columns = df.columns.str.strip()

            # Verificar si las columnas requeridas están presentes
            required_columns = [
                'Fecha día incluido', 'Movimientos Coppel', 'Importe Coppel', 'Movimientos Bancoppel', 'Importe Bancoppel'
            ]
            if not all(col in df.columns for col in required_columns):
                logging.error(f"Columnas requeridas no encontradas en {filename}. Se omitirá este archivo.")
                continue

            # Convertir columna de fecha a datetime y corregir diferencias de días
            df['Fecha día incluido'] = pd.to_datetime(df['Fecha día incluido'], errors='coerce')
            for i in range(1, len(df)):
                diferencia = (df.loc[i, 'Fecha día incluido'] - df.loc[i-1, 'Fecha día incluido']).days
                if diferencia > 15:
                    df.loc[i, 'Fecha día incluido'] = df.loc[i-1, 'Fecha día incluido'] + timedelta(days=1)

            # Eliminar filas con fechas inválidas
            df = df.dropna(subset=['Fecha día incluido'])

            # Validar si no quedan filas válidas
            if df.empty:
                logging.warning(f"No se encontraron filas válidas en {filename} después de filtrar por fecha.")
                continue

            # Registrar fechas procesadas
            try:
                fechas_procesadas = pd.to_datetime(df['Fecha día incluido']).dt.strftime('%y-%m-%d').unique()
                logging.info(f"Fechas procesadas en {filename}: {','.join(fechas_procesadas)}")
            except (KeyError, AttributeError) as e:
                logging.error(f"Error al procesar las fechas en {filename}: {e}")
            
            # Crear columnas adicionales
            df['AÑO'] = df['Fecha día incluido'].dt.year
            df['MES'] = df['Fecha día incluido'].dt.month
            df['SERVICIO'] = 'CFE'
            
            # Convertir columnas numéricas y manejar errores
            try:
                df['Movimientos Coppel'] = pd.to_numeric(df['Movimientos Coppel'], errors='coerce').fillna(0)
                df['Importe Coppel'] = pd.to_numeric(df['Importe Coppel'], errors='coerce').fillna(0)
                df['Movimientos Bancoppel'] = pd.to_numeric(df['Movimientos Bancoppel'], errors='coerce').fillna(0)
                df['Importe Bancoppel'] = pd.to_numeric(df['Importe Bancoppel'], errors='coerce').fillna(0)
            except ValueError:
                # Manejar columnas faltantes con valores por defecto
                df['Movimientos Coppel'] = df['Importe Coppel'] = 0
                df['Movimientos Bancoppel'] = df['Importe Bancoppel'] = 0

            # Sumar movimientos e importes
            df['Movimientos Coppel'] += df['Movimientos Bancoppel']
            df['Importe Coppel'] += df['Importe Bancoppel']

            # Seleccionar y renombrar columnas relevantes
            df = df[['AÑO', 'MES', 'SERVICIO', 'Movimientos Coppel', 'Importe Coppel', 'Movimientos Bancoppel', 'Importe Bancoppel']]
            df = df.rename(columns={'Movimientos Coppel': 'MOVIMIENTOS', 'Importe Coppel': 'IMPORTE'})

            # Agregar DataFrame a la lista
            dataframes.append(df)
        
        except Exception as e:
            logging.error(f"Error al procesar el archivo {filename}: {e}")
            continue

# Concatenar todos los DataFrames en uno solo
if dataframes:
    combined_df = pd.concat(dataframes, ignore_index=True)

    # Agrupar por AÑO, MES y SERVICIO, sumando movimientos e importes
    summary_df = combined_df.groupby(['AÑO', 'MES', 'SERVICIO'], as_index=False).agg({
        'MOVIMIENTOS': 'sum',
        'IMPORTE': 'sum'
    })

    # Formatear la columna 'IMPORTE' a formato de moneda
    summary_df['IMPORTE'] = summary_df['IMPORTE'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")

    # Guardar el resumen en un archivo Excel
    output_file = 'REPORTE_CONSOLIDADO_CFE.xlsx'
    summary_df.to_excel(output_file, index=False)

    logging.info(f"Todos los registros se han combinado y resumido exitosamente en '{output_file}'.")
    print(f"El archivo consolidado se ha guardado como '{output_file}'.")

    # Mostrar mensaje de éxito
    messagebox.showinfo("Éxito", "El proceso se completó exitosamente.")
else:
    # Mostrar mensaje de error si no se procesaron archivos
    logging.warning("No se procesaron archivos Excel debido a errores.")
    print("No se procesaron archivos Excel debido a errores.")
    messagebox.showerror("Error", "No se procesaron archivos Excel debido a errores.")