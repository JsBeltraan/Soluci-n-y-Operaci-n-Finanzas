import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine, text
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QDateEdit, QHBoxLayout, QSpacerItem, QSizePolicy, QDesktopWidget, QDialog, QTextEdit
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QDate
import sys
import traceback

# Clase personalizada para botón
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

# Clase personalizada para el botón con imagen
class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

# Ventana principal de la aplicación
class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte de Inspecciones de STPS")  # Título de la ventana
        self.setMinimumSize(600, 400)  # Tamaño mínimo de la ventana
        self.center_window()  # Centra la ventana en la pantalla

        # Layout principal vertical
        self.main_layout = QVBoxLayout(self)

        # Layout horizontal para el logo y el botón regresar
        self.header_layout = QHBoxLayout()

        # Label para el logo
        self.logo_label = QLabel(self)
        self.load_logo()  # Cargar el logo
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para empujar el botón de "Regresar" hacia la derecha
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)  # Conectar el evento para cerrar la ventana
        self.header_layout.addWidget(self.btn_salir)

        # Botón para abrir la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))  # Asignar el ícono al botón
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)  # Mostrar bitácora al hacer clic
        self.header_layout.addWidget(self.btn_bitacora)

        # Añadir el layout del header al layout principal
        self.main_layout.addLayout(self.header_layout)

        # Añadir un espaciador para empujar las fechas al centro
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Layout para las fechas de inicio y fin
        self.date_layout = QVBoxLayout()

        # Fecha inicio
        self.fecha_inicio_label = QLabel("Fecha de inicio:", self)
        self.date_layout.addWidget(self.fecha_inicio_label, alignment=Qt.AlignCenter)

        self.fecha_inicio_entry = QDateEdit(self)
        self.fecha_inicio_entry.setCalendarPopup(True)
        self.fecha_inicio_entry.setDate(QDate(datetime.now().year, 1, 1))  # Fecha por defecto (1 enero)
        self.date_layout.addWidget(self.fecha_inicio_entry, alignment=Qt.AlignCenter)

        # Fecha fin
        self.fecha_fin_label = QLabel("Fecha de fin:", self)
        self.date_layout.addWidget(self.fecha_fin_label, alignment=Qt.AlignCenter)

        self.fecha_fin_entry = QDateEdit(self)
        self.fecha_fin_entry.setCalendarPopup(True)
        self.fecha_fin_entry.setDate(QDate.currentDate())  # Fecha actual por defecto
        self.date_layout.addWidget(self.fecha_fin_entry, alignment=Qt.AlignCenter)

        # Añadir el layout de las fechas al layout principal
        self.main_layout.addLayout(self.date_layout)

        # Añadir otro espaciador para empujar el botón hacia abajo
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Botón para generar el reporte
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)  # Conectar el evento para generar reporte
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Información sobre el servidor y la base de datos
        self.SERVER = "SERVIDOR"
        self.BD = "BASEDEDATOS"
        self.USER = "USUARIO"
        self.PWD = "CONTRASEÑA"
        self.PORT = "PUERTO"

        # Consulta SQL para el reporte
        self.sql_query = """
        CONSULTA
        """

        # Cargar estilos desde un archivo CSS
        self.load_css()

    def load_css(self):
        """Carga el archivo CSS para aplicar estilos a los widgets."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())  # Aplica el CSS al widget
        else:
            print("El archivo style.css no se encontró.")

    def load_logo(self):
        """Cargar el logo de la empresa en la ventana."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)  # Asigna el logo al label
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def mostrar_bitacora(self):
        """Mostrar el cuadro de diálogo de la bitácora centrado en la pantalla."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)

        # Obtener la resolución de la pantalla y centrar el diálogo
        screen_geometry = QApplication.primaryScreen().geometry()

        dialog.show()

        layout = QVBoxLayout()

        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)  # Hacer que el texto sea solo de lectura
        bitacora_text.setPlainText(
            f"{self.SERVER}\n"
            f"{self.BD}\n\n"
            "Consulta ejecutada:\n"
            f"{self.sql_query}"
        )
        layout.addWidget(bitacora_text)

        dialog.setLayout(layout)
        dialog.exec_()

    def generar_reporte(self):
        """Función para generar el reporte de Excel según las fechas seleccionadas."""
        print("Inicia reporte")
        fecha_inicio = self.fecha_inicio_entry.date().toString("yyyy-MM-dd")  # Obtener fecha de inicio
        fecha_fin = self.fecha_fin_entry.date().toString("yyyy-MM-dd")  # Obtener fecha de fin

        # Definir la consulta SQL con las fechas
        sql = self.sql_query.replace("fecha_inicio", f"'{fecha_inicio}'").replace("fecha_fin", f"'{fecha_fin}'")

        # Crear la conexión SQLAlchemy
        engine = create_engine(
            f'postgresql://{self.USER}:{self.PWD}@{self.SERVER}:{self.PORT}/{self.BD}?client_encoding=utf-8'
        )

        try:
            # Establece una conexión y ejecuta la consulta SQL
            with engine.connect() as connection:
                result = connection.execute(text(sql))
                df = pd.DataFrame(result.fetchall(), columns=result.keys())  # Convertir el resultado en DataFrame
            
            # Crear el nombre del archivo Excel
            nombre_archivo = f'Reporte Inspecciones de STPS de {fecha_inicio} a {fecha_fin}.xlsx'

            # Mostrar el cuadro de diálogo para guardar el archivo
            options = QFileDialog.Options()
            ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options)

            if ruta_archivo:
                # Usar XlsxWriter como motor para guardar el archivo Excel
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Concentrado', index=False)

                    workbook = writer.book
                    worksheet = writer.sheets['Concentrado']

                    # Ajustar el ancho de las columnas
                    for col_num, column in enumerate(df.columns):
                        column_data = df[column].astype(str)
                        if pd.api.types.is_string_dtype(column_data):
                            column_len = column_data.str.len().max()
                        else:
                            column_len = column_data.apply(lambda x: len(str(x))).max()

                        column_len = max(column_len, len(column)) + 2
                        worksheet.set_column(col_num, col_num, column_len)

                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")  # Mensaje de éxito
        except Exception as e:
            # Si ocurre un error, se muestra un mensaje con los detalles del error
            error_message = f"Falló la conexión. Error: {str(e)} \n{traceback.format_exc()}"
            QMessageBox.critical(self, "Error de conexión", error_message)
        finally:
            engine.dispose()  # Liberar la conexión a la base de datos

    def center_window(self):
        """Centrar la ventana principal en la pantalla."""
        screen_geometry = QDesktopWidget().screenGeometry()
        window_geometry = self.geometry()
        self.move(
            (screen_geometry.width() - window_geometry.width()) // 2,
            (screen_geometry.height() - window_geometry.height()) // 2
        )

if __name__ == '__main__':
    app = QApplication(sys.argv)  # Iniciar la aplicación
    window = ReportWindow()
    window.show()  # Mostrar la ventana
    sys.exit(app.exec_())  # Ejecutar la aplicación
