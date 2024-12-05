import pandas as pd
from tkinter import Tk, filedialog, messagebox

def seleccionar_archivo(tipo_archivo):
    root = Tk()
    root.withdraw()
    messagebox.showwarning("Alerta", f"Debe seleccionar el {tipo_archivo}.")
    ruta_archivo = filedialog.askopenfilename(
        title=f"Seleccionar {tipo_archivo}",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
    )
    return ruta_archivo

def cargar_catalogo(ruta_catalogo):
    catalogo = pd.read_excel(ruta_catalogo)
    return catalogo

def cargar_datos(ruta_archivo):
    datos = pd.read_excel(ruta_archivo)
    return datos

def generar_txt(catalogo, datos, ruta_salida):
    with open(ruta_salida, 'w') as archivo_salida:
        for _, fila_datos in datos.iterrows():
            linea = ""
            for _, fila_cat in catalogo.iterrows():
                nombre_columna = fila_cat['Nombre de la cabecera del archivo a cargar']
                
                # Limpiar y convertir 'Espacios máximos' a entero
                espacios_maximos_str = str(fila_cat['Espacios máximos']).strip()
                espacios_maximos = int(espacios_maximos_str) if espacios_maximos_str.isdigit() else 0

                # Verifica si el campo existe en los datos y no es NaN
                if pd.isna(nombre_columna) or nombre_columna not in datos.columns:
                    # Rellenar con espacios en blanco si el campo no está presente
                    campo = " " * espacios_maximos
                else:
                    # Obtener el valor de la columna; si es NaN, asignar una cadena vacía
                    valor = fila_datos.get(nombre_columna, "")
                    if pd.isna(valor):
                        valor = ""

                    # Completar con ceros a la izquierda si es numérico
                    valor = str(valor)
                    if valor.isdigit():
                        campo = valor.zfill(espacios_maximos)
                    else:
                        campo = valor[:espacios_maximos].ljust(espacios_maximos)

                # Agregar el campo en orden
                linea += campo

            # Asegurar que la línea sea de exactamente 500 caracteres
            linea = linea[:500].ljust(500)

            archivo_salida.write(linea + "\n")

        # Agregar la línea extra al final del archivo - LINEA QUE PUEDE PEDIR USUARIO MODIFICACION
        linea_final = (
            "TLR" +                      # Posición 1-3: "TLR"
            "0".zfill(15) +              # Posición 4-18: 15 ceros
            "0".zfill(15) +              # Posición 19-33: 15 ceros
            "048".zfill(15) +            # Posición 34-48: "048" con ceros a la izquierda para completar 15
            "048".zfill(15) +            # Posición 49-63: "048" con ceros a la izquierda para completar 15
            " " * 37                     # Posición 64-100: 37 espacios en blanco
        )

        # Asegurar que la línea final tenga exactamente 101 caracteres
        linea_final = linea_final[:101].ljust(101)
        archivo_salida.write(linea_final + "\n")

# Selección de archivos
ruta_catalogo = seleccionar_archivo("Catálogo")
ruta_archivo_datos = seleccionar_archivo("Archivo de datos")

# Cargar el catálogo y los datos
catalogo = cargar_catalogo(ruta_catalogo)
datos = cargar_datos(ruta_archivo_datos)

# Generar el archivo TXT
ruta_archivo_salida = "PL500.txt"
generar_txt(catalogo, datos, ruta_archivo_salida)

# Mensaje de éxito al finalizar la exportación
messagebox.showinfo("Éxito", "El archivo Layout PL500.txt fue generado correctamente.")
