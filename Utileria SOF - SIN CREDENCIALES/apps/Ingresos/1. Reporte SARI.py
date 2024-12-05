import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QDateEdit, QHBoxLayout, QSpacerItem, QSizePolicy, QDesktopWidget, QDialog, QTextEdit
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QDate
import sys
import numpy as np

# Botón personalizado
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

# Botón personalizado para incluir íconos
class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

# Ventana principal de la aplicación
class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte de SARI")
        self.setMinimumSize(600, 400)
        self.center_window()  # Centrar la ventana

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # Layout para encabezado (logo y botones)
        self.header_layout = QHBoxLayout()

        # Agregar logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para separar el logo del botón
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón "Salir"
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)  # Cerrar ventana al hacer clic
        self.header_layout.addWidget(self.btn_salir)

        # Botón para abrir la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)  # Cargar ícono
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)  # Mostrar bitácora
        self.header_layout.addWidget(self.btn_bitacora)

        # Agregar encabezado al layout principal
        self.main_layout.addLayout(self.header_layout)

        # Espaciador para centrar las fechas
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Layout para selección de fechas
        self.date_layout = QVBoxLayout()

        # Campo de fecha inicio
        self.fecha_inicio_label = QLabel("Fecha de inicio:", self)
        self.date_layout.addWidget(self.fecha_inicio_label, alignment=Qt.AlignCenter)

        self.fecha_inicio_entry = QDateEdit(self)
        self.fecha_inicio_entry.setCalendarPopup(True)  # Habilitar selector de calendario
        self.fecha_inicio_entry.setDate(QDate(datetime.now().year, datetime.now().month, 1))  # Fecha predeterminada
        self.date_layout.addWidget(self.fecha_inicio_entry, alignment=Qt.AlignCenter)

        # Campo de fecha fin
        self.fecha_fin_label = QLabel("Fecha de fin:", self)
        self.date_layout.addWidget(self.fecha_fin_label, alignment=Qt.AlignCenter)

        self.fecha_fin_entry = QDateEdit(self)
        self.fecha_fin_entry.setCalendarPopup(True)  # Habilitar selector de calendario
        self.fecha_fin_entry.setDate(QDate.currentDate())  # Fecha predeterminada
        self.date_layout.addWidget(self.fecha_fin_entry, alignment=Qt.AlignCenter)

        # Agregar fechas al layout principal
        self.main_layout.addLayout(self.date_layout)

        # Espaciador para mover el botón hacia abajo
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Botón para generar reporte
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)  # Conectar al método correspondiente
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Información del servidor y base de datos
        self.SERVER = "SERVIDOR"
        self.BD = "BASEDEDATOS"
        self.USER = "USUARIO"
        self.PWD = "CONTRASEÑA"
        self.PORT = "PUERTO"

        # Consulta SQL predeterminada
        self.sql_query = """
        CONSULTA    
        """

        # Cargar estilos desde archivo CSS
        self.load_css()

    def load_css(self):
        """Cargar el archivo CSS para estilos."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())
        else:
            print("El archivo style.css no se encontró.")

    def load_logo(self):
        """Cargar y mostrar el logo."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def mostrar_bitacora(self):
        """Abrir una ventana con la información de la bitácora."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)

        layout = QVBoxLayout()

        # Texto de la bitácora
        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)  # Campo no editable
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
        """Generar reporte con los datos ingresados."""
        print("Inicia reporte")
        fecha_inicio = self.fecha_inicio_entry.date().toString("yyyy-MM-dd")
        fecha_fin = self.fecha_fin_entry.date().toString("yyyy-MM-dd")

        # Reemplazar fechas en la consulta
        sql = self.sql_query.replace("fecha_inicio", f"'{fecha_inicio}'").replace("fecha_fin", f"'{fecha_fin}'")

        # Crear conexión con la base de datos
        engine = create_engine(
            f'postgresql://{self.USER}:{self.PWD}@{self.SERVER}:{self.PORT}/{self.BD}?client_encoding=utf-8'
        )

        try:
            # Ejecutar consulta y guardar resultados en DataFrame
            df = pd.read_sql_query(sql, engine)
            nombre_archivo = f'Reporte Sari - {fecha_inicio} a {fecha_fin}.xlsx'

            # Diálogo para seleccionar ruta de guardado
            options = QFileDialog.Options()
            ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options)

            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    # Reemplazar valores no válidos
                    df = df.replace([np.nan, np.inf, -np.inf], '')

                    # Escribir datos en Excel
                    df.to_excel(writer, sheet_name='Concentrado', index=False, startrow=1, header=False)

                    # Aplicar formato y ajustar ancho de columnas
                    workbook = writer.book
                    worksheet = writer.sheets['Concentrado']

                    # Formato para encabezados
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'center',
                        'align': 'center',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })

                    for col_num, column in enumerate(df.columns):
                        worksheet.write(0, col_num, column, header_format)
                        column_len = max(df[column].astype(str).str.len().max(), len(column)) + 2
                        worksheet.set_column(col_num, col_num, column_len)

                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")
        except Exception as e:
            error_message = f"Falló la conexión. Error: {str(e)}"
            QMessageBox.critical(self, "Error de conexión", error_message)

    def center_window(self):
        """Centrar ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

# Ejecutar la aplicación
if __name__ == "__main__":
    app = QApplication([])
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())
