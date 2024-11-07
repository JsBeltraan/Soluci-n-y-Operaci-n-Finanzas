
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
from tkinter import filedialog
from sqlalchemy import create_engine, text
import pandas as pd
from datetime import datetime, timedelta

# Conexión a la base de datos PostgreSQL
engine = create_engine('postgresql://USER:PASS@00.00.00.00:0000/BD')

# Función para obtener el primer y último día del mes anterior
def fechas_mes_anterior():
    hoy = datetime.today()
    primer_dia_mes_actual = hoy.replace(day=1)
    ultimo_dia_mes_anterior = primer_dia_mes_actual - timedelta(days=1)
    primer_dia_mes_anterior = ultimo_dia_mes_anterior.replace(day=1)
    return primer_dia_mes_anterior, ultimo_dia_mes_anterior

# Función para obtener datos desde la base de datos
def obtener_reporte(id_producto, fecha_inicio=None, fecha_fin=None, por_mes=True):
    query = """
        CONSULTA
    """
    if por_mes:
        query += " + QUERY"
    elif fecha_inicio and fecha_fin:
        query += " + QUERY"
    query += """ + QUERY"""

    # Mostrar el query en la bitácora
    log_text.set(f"Query ejecutado:\n{query}\nCon id_producto={id_producto}, fecha_inicio={fecha_inicio}, fecha_fin={fecha_fin}")
    try:
        with engine.connect() as connection:
            result = connection.execute(text(query), {'id_producto': id_producto, 'fecha_inicio': fecha_inicio, 'fecha_fin': fecha_fin})
            df = pd.DataFrame(result.fetchall(), columns=result.keys())
            #print(df)
            tipo = tipo_reporte.get()
            nombre_archivo = f'Diferencias de {tipo}.xlsx'

            # Muestra el cuadro de diálogo para guardar el archivo y obtiene la ruta del archivo
            ruta_archivo = filedialog.asksaveasfilename(
                initialfile=nombre_archivo, defaultextension=".xlsx", filetypes=[
                    ("Excel files", "*.xlsx")])

            # if ruta_archivo: 
            # Crea un nuevo archivo de Excel
            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    
                    # Escribe el DataFrame completo en la primera hoja
                    df.to_excel(writer, sheet_name=f'{tipo}', index=False)

                    # Acceder al workbook y worksheet
                    workbook = writer.book
                    worksheet = writer.sheets[f'{tipo}']

                    # Ajusta el ancho de las columnas al tamaño de la información de cada columna
                    for sheet in writer.sheets:
                        worksheet = writer.sheets[sheet]
                        ##worksheet = writer.sheets['Pestamos']
                        #worksheet.insert_cols(0, 2)  # Inserta 2 columnas al inicio
                
                        # Inserta el texto en la celda A1:C2
                        #worksheet.merge_range('A1:C2', nombre_archivo)  
                        for col_num, column in enumerate(df.columns):
                            column_len = df[column].astype(str).str.len().max()
                            column_len = max(column_len, len(column)) + 2
                            worksheet.set_column(col_num, col_num, column_len)

            messagebox.showinfo(
                "Éxito",
                "El reporte fue generado con éxito.")
                
    except Exception as e:
    # Manejo de la excepción y mensaje de error
        error_message = f"Falló la conexión. Error: {str(e)}"
        messagebox.showinfo("Error de conexión", error_message)
# Función que se ejecuta al seleccionar el reporte y rango de fechas
def generar_reporte():
    tipo = tipo_reporte.get()
    if tipo == 'Transferencias'  :
        id_producto = '112' 
    elif tipo == 'Corresponsales':
        id_producto = '056'
    
    fecha_inicio = fecha_inicio_input.get_date() if rango_fecha.get() else None
    fecha_fin = fecha_fin_input.get_date() if rango_fecha.get() else None
    por_mes = opcion_periodo.get() == 'Por mes'

    obtener_reporte(id_producto, fecha_inicio, fecha_fin, por_mes)

# Obtener fechas del mes anterior
fecha_inicio_default, fecha_fin_default = fechas_mes_anterior()

# Crear la ventana principal
root = tk.Tk()
root.title("Seleccionar Reporte")

# Etiquetas y campos de selección
ttk.Label(root, text="Seleccione el Tipo de Reporte:").grid(column=0, row=0, padx=10, pady=10)
tipo_reporte = tk.StringVar(value="Transferencias")  # Valor por defecto
ttk.Combobox(root, textvariable=tipo_reporte, values=["Transferencias", "Corresponsales"]).grid(column=1, row=0, padx=10, pady=10)

ttk.Label(root, text="Seleccione el Periodo:").grid(column=0, row=1, padx=10, pady=10)
opcion_periodo = tk.StringVar(value="Por mes")  # Valor por defecto
ttk.Combobox(root, textvariable=opcion_periodo, values=["Por mes", "Por rango de fechas"]).grid(column=1, row=1, padx=10, pady=10)

# Opciones para el rango de fechas
rango_fecha = tk.BooleanVar()
rango_fecha.set(False)
tk.Checkbutton(root, text="¿Usar rango de fechas?", variable=rango_fecha).grid(column=0, row=2, padx=10, pady=10)

ttk.Label(root, text="Fecha inicio:").grid(column=0, row=3, padx=10, pady=10)
fecha_inicio_input = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, year=fecha_inicio_default.year, month=fecha_inicio_default.month, day=fecha_inicio_default.day)
fecha_inicio_input.grid(column=1, row=3, padx=10, pady=10)

ttk.Label(root, text="Fecha fin:").grid(column=0, row=4, padx=10, pady=10)
fecha_fin_input = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, year=fecha_fin_default.year, month=fecha_fin_default.month, day=fecha_fin_default.day)
fecha_fin_input.grid(column=1, row=4, padx=10, pady=10)

# Botón para generar el reporte
ttk.Button(root, text="Generar Reporte", command=generar_reporte).grid(column=0, row=5, columnspan=2, padx=10, pady=10)

# Bitácora para mostrar el query ejecutado
log_text = tk.StringVar()
log_label = tk.Label(root, textvariable=log_text, justify="left", anchor="w", relief="sunken", bg="white", width=80, height=5)
log_label.grid(column=0, row=6, columnspan=2, padx=10, pady=10)

# Iniciar la aplicación
root.mainloop()
