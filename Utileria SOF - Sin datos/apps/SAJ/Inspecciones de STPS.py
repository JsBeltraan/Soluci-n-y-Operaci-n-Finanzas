import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine, text
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QDateEdit, QHBoxLayout, QSpacerItem, QSizePolicy, QDesktopWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QDate
import sys
import traceback

class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte de Inspecciones de STPS")
        self.setMinimumSize(600, 400)
        self.center_window()


        # Layout principal vertical
        self.main_layout = QVBoxLayout(self)

        # Layout horizontal para el logo y el botón regresar
        self.header_layout = QHBoxLayout()

        # Label para el logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Añadir el layout del header al layout principal
        self.main_layout.addLayout(self.header_layout)

        # Añadir un espaciador para empujar las fechas al centro
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Layout para las fechas
        self.date_layout = QVBoxLayout()

        # Fecha inicio
        self.fecha_inicio_label = QLabel("Fecha de inicio:", self)
        self.date_layout.addWidget(self.fecha_inicio_label, alignment=Qt.AlignCenter)

        self.fecha_inicio_entry = QDateEdit(self)
        self.fecha_inicio_entry.setCalendarPopup(True)
        self.fecha_inicio_entry.setDate(QDate(datetime.now().year, 1, 1))
        self.date_layout.addWidget(self.fecha_inicio_entry, alignment=Qt.AlignCenter)

        # Fecha fin
        self.fecha_fin_label = QLabel("Fecha de fin:", self)
        self.date_layout.addWidget(self.fecha_fin_label, alignment=Qt.AlignCenter)

        self.fecha_fin_entry = QDateEdit(self)
        self.fecha_fin_entry.setCalendarPopup(True)
        self.fecha_fin_entry.setDate(QDate.currentDate())
        self.date_layout.addWidget(self.fecha_fin_entry, alignment=Qt.AlignCenter)

        # Añadir el layout de las fechas al layout principal
        self.main_layout.addLayout(self.date_layout)

        # Añadir otro espaciador para empujar el botón hacia abajo
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Botón para generar reporte al final
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Aplicar estilos desde archivo CSS
        self.load_css()

    def load_css(self):
        """Carga el archivo CSS para aplicar estilos a los widgets."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())
        else:
            print("El archivo style.css no se encontró.")

    def load_logo(self):
        """Cargar el logo."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def generar_reporte(self):
        """Función para generar el reporte de Excel."""
        print("Inicia reporte")
        fecha_inicio = self.fecha_inicio_entry.date().toString("yyyy-MM-dd")
        fecha_fin = self.fecha_fin_entry.date().toString("yyyy-MM-dd")

        # Define la consulta SQL
        sql = f"""
        CONSULTA        """

        # Crea una conexión SQLAlchemy
        engine = create_engine(
            'postgresql://USER:PASS@00.00.00.00:0000/BD?client_encoding=utf8'
        )

        try:
        # Establece una conexión explícita
            with engine.connect() as connection:
                # Ejecuta la consulta SQL usando 'text()'
                result = connection.execute(text(sql))
                # Conviértelo en un DataFrame
                df = pd.DataFrame(result.fetchall(), columns=result.keys())
            
            nombre_archivo = f'Reporte Inspecciones de STPS de {fecha_inicio} a {fecha_fin}.xlsx'

            # Muestra el cuadro de diálogo para guardar el archivo
            options = QFileDialog.Options()
            ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options)

            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Concentrado', index=False)

                    # Obtén el workbook y worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Concentrado']

                    # Ajustar el ancho de las columnas
                    for col_num, column in enumerate(df.columns):
                        # Convertimos la columna a cadenas
                        column_data = df[column].astype(str)
                        
                        # Verificamos si la columna tiene el atributo 'str' y calculamos su longitud
                        if pd.api.types.is_string_dtype(column_data):
                            column_len = column_data.str.len().max()
                        else:
                            # Si no es de tipo string, calculamos la longitud como una cadena
                            column_len = column_data.apply(lambda x: len(str(x))).max()

                        # Establecemos un mínimo tamaño de columna para evitar que sea muy pequeño
                        column_len = max(column_len, len(column)) + 2
                        worksheet.set_column(col_num, col_num, column_len)


                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")
        except Exception as e:
            error_message = f"Falló la conexión. Error: {str(e)} \n{traceback.format_exc()}"
            QMessageBox.critical(self, "Error de conexión", error_message)
        finally:
            engine.dispose()

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry() # Obtener dimensiones de la ventana
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

if __name__ == "__main__":
    app = QApplication([])
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())
