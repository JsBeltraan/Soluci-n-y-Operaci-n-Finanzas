import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QDateEdit, QHBoxLayout, QSpacerItem, QSizePolicy, QDesktopWidget, QDialog, QTextEdit
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QDate
import sys
import numpy as np

class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte de SARI")
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

        # Espaciador para empujar el botón de "Regresar" hacia la derecha
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)
        self.header_layout.addWidget(self.btn_salir)

        # Botón para abrir la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.header_layout.addWidget(self.btn_bitacora)

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
        self.fecha_inicio_entry.setDate(QDate(datetime.now().year, datetime.now().month, 1))
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

        self.setLayout(self.main_layout)

        # Botón para generar reporte al final
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Información sobre servidor, base de datos
        self.SERVER = "SERVER"
        self.BD = "BD"
        self.USER = "USER"
        self.PWD = "PWD"
        self.PORT = "PORT"

        # Consulta SQL para la bitácora
        self.sql_query = """
        CONSULTA1
            """

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

    def mostrar_bitacora(self):
        """Mostrar el cuadro de diálogo de la bitácora centrado en la pantalla."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)

        # Obtener la resolución de la pantalla
        screen_geometry = QApplication.primaryScreen().geometry()

        dialog.show()

        layout = QVBoxLayout()

        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)
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
        """Función para generar el reporte de Excel."""
        print("Inicia reporte")
        fecha_inicio = self.fecha_inicio_entry.date().toString("yyyy-MM-dd")
        fecha_fin = self.fecha_fin_entry.date().toString("yyyy-MM-dd")

        # Define la consulta SQL
        sql = self.sql_query.replace("fecha_inicio", f"'{fecha_inicio}'").replace("fecha_fin", f"'{fecha_fin}'")

        # Crea conexión
        engine = create_engine(
            f'postgresql://{self.USER}:{self.PWD}@{self.SERVER}:{self.PORT}/{self.BD}?client_encoding=utf-8'
        )

        try:
            # Ejecutar consulta y guardar los resultados en un df
            df = pd.read_sql_query(sql, engine)
            nombre_archivo = f'Reporte Sari - {fecha_inicio} a {fecha_fin}.xlsx'

            # Mostrar cuadro de diálogo para guardar el archivo
            options = QFileDialog.Options()
            ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options)

            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    # Reemplazar valores NaN e Inf antes de escribir al archivo
                    df = df.replace([np.nan, np.inf, -np.inf], '')

                    # Escribe el DataFrame completo en la primera hoja
                    df.to_excel(writer, sheet_name='Concentrado', index=False, startrow=1, header=False)

                    # Obtener el workbook y worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Concentrado']

                    # Formato para los encabezados
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'center',
                        'align': 'center',
                        'fg_color': '#D7E4BC',  # Color de fondo para los encabezados
                        'border': 1  # Bordes en los encabezados
                    })

                    # Escribir los encabezados del DataFrame manualmente con el formato
                    for col_num, column in enumerate(df.columns):
                        worksheet.write(0, col_num, column, header_format)

                    # Crear un formato para las celdas con bordes
                    cell_format = workbook.add_format({
                        'border': 1,
                        'align': 'left',
                        'valign': 'vcenter'
                    })

                    # Ajustar el ancho de las columnas al tamaño de la información
                    for col_num, column in enumerate(df.columns):
                        column_len = df[column].astype(str).str.len().max()
                        column_len = max(column_len, len(column)) + 2
                        worksheet.set_column(col_num, col_num, column_len)

                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")
        except Exception as e:
            error_message = f"Falló la conexión. Error: {str(e)}"
            QMessageBox.critical(self, "Error de conexión", error_message)

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()  # Obtener dimensiones de la ventana
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

if __name__ == "__main__":
    app = QApplication([])
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())
