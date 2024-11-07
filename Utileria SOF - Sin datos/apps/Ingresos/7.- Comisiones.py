import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry
import pandas as pd
from sqlalchemy import create_engine

def generar_reporte():
    fecha_inicio = fecha_inicio_entry.get_date()
    fecha_fin = fecha_fin_entry.get_date()
    # Define la consulta SQL
    sql = f"""
    CONSULTA
    """
    
    try:
        # Crea una conexión
        engine = create_engine(
            'postgresql://USER:PASS@00.00.00.00:0000/BD')
        
        
        # Ejecuta la consulta y guarda los resultados en un DataFrame
        df = pd.read_sql_query(sql, engine)

        nombre_archivo = f'Reporte comisiones DYA del {fecha_inicio} al {fecha_fin}.xlsx'

        # Muestra el cuadro de diálogo para guardar el archivo y obtiene la ruta del archivo
        ruta_archivo = filedialog.asksaveasfilename(
            initialfile=nombre_archivo, defaultextension=".xlsx", filetypes=[
                ("Excel files", "*.xlsx")])

        # if ruta_archivo:
        # Crea un nuevo archivo de Excel
        if ruta_archivo:
            with pd.ExcelWriter(ruta_archivo) as writer:
                # Escribe el DataFrame completo en la primera hoja
                df.to_excel(writer, sheet_name='Concentrado', index=False)
                # Calcula el total de la columna 'cantidad'
                total_concentrado = df['comision'].sum()
                # Escribe el total en la primera fila a la derecha del DataFrame
                workbook = writer.book
                worksheet = writer.sheets['Concentrado']
                worksheet.write(0, len(df.columns) + 1, 'Total')
                worksheet.write(1, len(df.columns) + 1, total_concentrado)

                # Para cada fecha de movimiento en el DataFrame
                for fecha in df['envio'].unique():
                    # Filtra el DataFrame para solo esa fecha
                    df_fecha = df[df['envio'] == fecha]
                    # Calcula el total de la columna 'cantidad'
                    total = df_fecha['comision'].sum()
                    # Escribe el DataFrame de esa fecha en una hoja de Excel
                    df_fecha.to_excel(writer, sheet_name=f'{fecha}', index=False)
                    # Escribe el total en la primera fila a la derecha del DataFrame
                    worksheet_fecha = writer.sheets[f'{fecha}']
                    worksheet_fecha.write(0, len(df_fecha.columns) + 1, 'Total')
                    worksheet_fecha.write(1, len(df_fecha.columns) + 1, total)
        

                # Ajusta el ancho de las columnas al tamaño de la información de cada columna
                for sheet in writer.sheets:
                    worksheet = writer.sheets[sheet]
                    for col_num, column in enumerate(df.columns):
                        column_len = df[column].astype(str).str.len().max()
                        column_len = max(column_len, len(column)) + 2
                        worksheet.set_column(col_num, col_num, column_len)

        messagebox.showinfo(
            "Éxito",
            "El reporte fue generado con éxito.")  # Mensaje de éxito
            
    except Exception as e:
    # Manejo de la excepción y mensaje de error
        error_message = f"Falló la conexión. Error: {str(e)}"
        messagebox.showinfo("Error de conexión", error_message)
    
class BotonRedondeado(tk.Button):
    def __init__(self, master=None, **kwargs):
        tk.Button.__init__(self, master, **kwargs)
        self.config(relief=tk.FLAT, bg="#A51427", fg="white", activebackground="#D62E4D", activeforeground="white", cursor="hand2")
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, event):
        self.config(bg="#D62E4D")

    def on_leave(self, event):
        self.config(bg="#A51427")

root = tk.Tk()
root.geometry("600x400")
root.title("Reporte de Comisiones DYA")
root.configure(bg="#003366")

fecha_inicio_label = tk.Label(root, text="Fecha de inicio:", font=("Arial", 14), bg="#003366", fg="white")
fecha_inicio_entry = DateEntry(root, font=("Arial", 12), width=12, background="white", foreground="black",date_pattern = 'DD/MM/YYYY')
fecha_inicio_label.pack(pady=10)
fecha_inicio_entry.pack(pady=5)

fecha_fin_label = tk.Label(root, text="Fecha de fin:", font=("Arial", 14), bg="#003366", fg="white")
fecha_fin_entry = DateEntry(root, font=("Arial", 12), width=12, background="white", foreground="black",date_pattern = 'DD/MM/YYYY')
fecha_fin_label.pack(pady=10)
fecha_fin_entry.pack(pady=5)

generar_reporte_button = BotonRedondeado(root, text="Generar reporte", command=generar_reporte, font=("Arial", 14), width=20, padx=20, pady=10)
generar_reporte_button.pack(pady=20)

root.mainloop()