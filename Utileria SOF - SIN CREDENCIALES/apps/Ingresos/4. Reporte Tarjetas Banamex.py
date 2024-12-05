import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QDateEdit, QHBoxLayout, QSpacerItem, QSizePolicy, QDesktopWidget, QDialog, QTextEdit
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QDate
import sys

# Clase personalizada para un botón
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

# Clase personalizada para un botón que incluye una imagen
class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

# Clase principal que representa la ventana del reporte
class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte Tarjetas Banamex")  # Título de la ventana
        self.setMinimumSize(600, 400)  # Tamaño mínimo de la ventana
        self.center_window()  # Centra la ventana en la pantalla

        # Layout principal y layouts secundarios
        self.main_layout = QVBoxLayout(self)
        self.header_layout = QHBoxLayout()  # Layout para el logo y botones
        self.date_layout = QVBoxLayout()  # Layout para seleccionar fechas

        # Carga del logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para alinear botones
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)  # Cierra la aplicación
        self.header_layout.addWidget(self.btn_salir)

        # Botón para mostrar bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.header_layout.addWidget(self.btn_bitacora)

        # Agrega el header y espaciadores al layout principal
        self.main_layout.addLayout(self.header_layout)
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Configuración de widgets para selección de fechas
        self.fecha_inicio_label = QLabel("Fecha de inicio:", self)
        self.date_layout.addWidget(self.fecha_inicio_label, alignment=Qt.AlignCenter)
        self.fecha_inicio_entry = QDateEdit(self)
        self.fecha_inicio_entry.setCalendarPopup(True)
        self.fecha_inicio_entry.setDate(QDate(datetime.now().year, datetime.now().month, 1))  # Fecha predeterminada
        self.date_layout.addWidget(self.fecha_inicio_entry, alignment=Qt.AlignCenter)
        self.fecha_fin_label = QLabel("Fecha de fin:", self)
        self.date_layout.addWidget(self.fecha_fin_label, alignment=Qt.AlignCenter)
        self.fecha_fin_entry = QDateEdit(self)
        self.fecha_fin_entry.setCalendarPopup(True)
        self.fecha_fin_entry.setDate(QDate.currentDate())
        self.date_layout.addWidget(self.fecha_fin_entry, alignment=Qt.AlignCenter)
        self.main_layout.addLayout(self.date_layout)

        # Botón para generar el reporte
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Información de conexión a la base de datos y consulta SQL
        self.SERVER = "SERVIDOR"
        self.BD = "BASEDEDATOS"
        self.USER = "USUARIO"
        self.PWD = "PASSWORD"
        self.PORT = "PORT"
        self.sql_query = "CONSULTA"

        # Aplicar estilos desde un archivo CSS
        self.load_css()

    def load_css(self):
        """Carga un archivo CSS para los estilos de la interfaz."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())

    def load_logo(self):
        """Carga el logo en el widget correspondiente."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)

    def mostrar_bitacora(self):
        """Abre un cuadro de diálogo mostrando la bitácora."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)
        layout = QVBoxLayout()
        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)
        bitacora_text.setPlainText(f"{self.SERVER}\n{self.BD}\n\nConsultas ejecutadas:\n{self.sql_query}")
        layout.addWidget(bitacora_text)
        dialog.setLayout(layout)
        dialog.exec_()

    def generar_reporte(self):
        """Genera un reporte de Excel basado en las fechas seleccionadas."""
        fecha_inicio = self.fecha_inicio_entry.date().toString("yyyy-MM-dd")
        fecha_fin = self.fecha_fin_entry.date().toString("yyyy-MM-dd")
        sql = self.sql_query.replace("fecha_inicio", fecha_inicio).replace("fecha_fin", fecha_fin)
        engine = create_engine(f'postgresql://{self.USER}:{self.PWD}@{self.SERVER}:{self.PORT}/{self.BD}')
        try:
            df = pd.read_sql_query(sql, engine)
            nombre_archivo = f'Reporte Tarjetas Banamex - {fecha_inicio} a {fecha_fin}.xlsx'
            options = QFileDialog.Options()
            ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options)
            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Concentrado', index=False)
                    workbook = writer.book
                    worksheet = writer.sheets['Concentrado']
                    for col_num, column in enumerate(df.columns):
                        column_len = max(df[column].astype(str).str.len().max(), len(column)) + 5
                        worksheet.set_column(col_num, col_num, column_len)
                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")
        except Exception as e:
            QMessageBox.critical(self, "Error de conexión", f"Falló la conexión. Error: {str(e)}")

    def center_window(self):
        """Centra la ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

# Inicialización de la aplicación
if __name__ == "__main__":
    app = QApplication([])
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())
