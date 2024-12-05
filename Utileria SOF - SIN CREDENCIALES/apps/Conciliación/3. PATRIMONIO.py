import pandas as pd
import os
from tkinter import Tk, messagebox
from tkinter.filedialog import askdirectory

# Crear una ventana raíz de Tkinter (no se mostrará)
root = Tk()
root.withdraw()

# Alerta antes de seleccionar la carpeta
messagebox.showinfo("Seleccionar Carpeta", "Seleccionar carpeta con archivos xlsx u ods de PATRIMONIO")

# Abrir un diálogo para seleccionar la carpeta
folder_path = askdirectory(title='Seleccionar Carpeta con Archivos Excel')

# Inicializar una lista vacía para almacenar los DataFrames
dataframes = []

# Iterar sobre cada archivo en la carpeta
for filename in os.listdir(folder_path):  # Recorre todos los archivos en la carpeta seleccionada
    if filename.endswith('.xls') or filename.endswith('.xlsx'):  # Verifica que el archivo sea Excel
        # Construir la ruta completa del archivo
        file_path = os.path.join(folder_path, filename)
        
        try:
            # Leer las primeras 15 filas del archivo para detectar dónde comienza el layout
            temp_df = pd.read_excel(file_path, sheet_name='Reporte General', nrows=15, dtype=str)
            
            # Buscar la fila que contiene una fecha válida
            date_row = temp_df[temp_df.iloc[:, 0].str.match(r'\d{2}/\d{2}/\d{4}', na=False)].index[0]
            
            # Ahora leer todo el archivo desde la fila donde empieza el layout
            df = pd.read_excel(file_path, sheet_name='Reporte General', skiprows=date_row, dtype=str).replace("'", "")
            
            # Limpiar los nombres de las columnas (eliminar espacios)
            df.columns = df.columns.str.strip()
            
            # Verificar si las columnas 'FECHA', 'MOVIMIENTOS' y 'IMPORTE' existen en el DataFrame
            required_columns = ['FECHA', 'MOVIMIENTOS', 'IMPORTE']
            if not all(col in df.columns for col in required_columns):  # Valida que todas las columnas requeridas estén presentes
                print(f"Columnas requeridas no encontradas en {filename}. Se omitirán estos datos.")
                continue
            
            # Filtrar solo las filas que contienen una fecha válida
            df = df[df['FECHA'].str.match(r'\d{2}/\d{2}/\d{4}', na=False)]
            
            # Convertir la columna de fecha a tipo datetime
            df['Fecha'] = pd.to_datetime(df['FECHA'], format='%d/%m/%Y', errors='coerce')
            
            # Eliminar las filas donde la fecha no se pueda convertir (esto elimina los totales)
            df = df.dropna(subset=['Fecha'])
            
            # Extraer año y mes de la fecha
            df['AÑO'] = df['Fecha'].dt.year
            df['MES'] = df['Fecha'].dt.month
            
            # Crear la columna SERVICIO (siempre será Patrimonio)
            df['SERVICIO'] = 'PATRIMONIO'
            
            # Usar la columna 'MOVIMIENTOS' para las transacciones
            df['MOVIMIENTOS'] = pd.to_numeric(df['MOVIMIENTOS'], errors='coerce')
            
            # Limpiar la columna 'IMPORTE', eliminar símbolo de peso y comas, y convertir a numérico
            df['IMPORTE'] = df['IMPORTE'].replace({'\$': '', ',': ''}, regex=True)
            df['IMPORTE'] = pd.to_numeric(df['IMPORTE'], errors='coerce')
            
            # Seleccionar las columnas relevantes
            df = df[['AÑO', 'MES', 'SERVICIO', 'MOVIMIENTOS', 'IMPORTE']]
            
            # Agregar el DataFrame a la lista
            dataframes.append(df)
        
        except Exception as e:  # Manejo de errores en la lectura de cada archivo
            print(f"Error al procesar el archivo {filename}: {e}")
            continue

# Concatenar todos los DataFrames en uno solo
if dataframes:
    combined_df = pd.concat(dataframes, ignore_index=True)

    # Agrupar por AÑO, MES, SERVICIO y sumar los movimientos y el importe
    summary_df = combined_df.groupby(['AÑO', 'MES', 'SERVICIO'], as_index=False).agg({
        'MOVIMIENTOS': 'sum',
        'IMPORTE': 'sum'
    })

    # Formatear la columna 'IMPORTE' a pesos, agregando el símbolo de pesos y comas
    summary_df['IMPORTE'] = summary_df['IMPORTE'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")

    # Guardar el DataFrame combinado en un nuevo archivo de Excel
    output_file = 'REPORTE_CONSOLIDADO_PATRIMONIO.xlsx'
    summary_df.to_excel(output_file, index=False)

    print(f"Todos los registros se han combinado y resumido exitosamente en '{output_file}'.")

    # Alerta de éxito al finalizar
    messagebox.showinfo("Proceso Completado", f"Todos los registros se han combinado y resumido exitosamente en '{output_file}'.")

else:
    print("No se procesaron archivos Excel debido a errores.")
    messagebox.showwarning("Error", "No se procesaron archivos Excel debido a errores.")
