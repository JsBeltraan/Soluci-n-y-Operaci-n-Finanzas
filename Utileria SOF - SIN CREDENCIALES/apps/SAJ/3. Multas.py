import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine, text
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout,
    QHBoxLayout, QSpacerItem, QSizePolicy, QDateEdit, QTextEdit, QDesktopWidget, QDialog
)
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QDate
import sys

# Clase personalizada para los botones
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

# Clase personalizada para el botón de la imagen de la bitácora
class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

# Ventana principal para generar reportes de multas
class ReportMultasWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte de Multas")  # Título de la ventana
        self.setMinimumSize(600, 400)  # Tamaño mínimo de la ventana
        self.center_window()  # Centra la ventana en la pantalla

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # Layout del encabezado con logo y botones
        self.header_layout = QHBoxLayout()

        # Label para cargar el logo
        self.logo_label = QLabel(self)
        self.load_logo()  # Método para cargar el logo
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack("Salir", self)
        self.btn_salir.clicked.connect(self.close)
        self.header_layout.addWidget(self.btn_salir)

        # Botón para abrir la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap("img/bitacora.png").scaled(20, 20, Qt.KeepAspectRatio)
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.header_layout.addWidget(self.btn_bitacora)

        # Añadir layout del encabezado al layout principal
        self.main_layout.addLayout(self.header_layout)

        # Layout para las fechas de inicio y fin
        self.date_layout = QVBoxLayout()

        # Fecha de inicio
        self.fecha_inicio_label = QLabel("Fecha de inicio:", self)
        self.date_layout.addWidget(self.fecha_inicio_label, alignment=Qt.AlignCenter)
        self.fecha_inicio_entry = QDateEdit(self)
        self.fecha_inicio_entry.setCalendarPopup(True)
        self.fecha_inicio_entry.setDate(QDate(datetime.now().year, 1, 1))  # Fecha predeterminada
        self.date_layout.addWidget(self.fecha_inicio_entry, alignment=Qt.AlignCenter)

        # Fecha de fin
        self.fecha_fin_label = QLabel("Fecha de fin:", self)
        self.date_layout.addWidget(self.fecha_fin_label, alignment=Qt.AlignCenter)
        self.fecha_fin_entry = QDateEdit(self)
        self.fecha_fin_entry.setCalendarPopup(True)
        self.fecha_fin_entry.setDate(QDate.currentDate())  # Fecha actual predeterminada
        self.date_layout.addWidget(self.fecha_fin_entry, alignment=Qt.AlignCenter)

        # Añadir layout de las fechas al layout principal
        self.main_layout.addLayout(self.date_layout)

        # Espaciador para empujar el botón hacia abajo
        self.main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Botón para generar reporte
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Datos de conexión a la base de datos
        self.SERVER = "SERVIDOR"
        self.BD = "BASEDEDATOS"
        self.USER = "USUARIO"
        self.PWD = "CONTRASEÑA"
        self.PORT = "PUERTO"
        self.sql_query = """
        CONSULTA
        """  # La consulta SQL que se ejecutará

        # Cargar el archivo de estilo CSS
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
        """Carga el logo desde el archivo especificado."""
        logo_path = "img/Logo-coppel.png"
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def mostrar_bitacora(self):
        """Muestra la bitácora con detalles de la conexión y consulta ejecutada."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)

        layout = QVBoxLayout(dialog)

        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)
        bitacora_text.setPlainText(
            f"Servidor: {self.SERVER}\n"
            f"Base de datos: {self.BD}\n\n"
            f"Consulta ejecutada:\n{self.sql_query}"
        )
        layout.addWidget(bitacora_text)

        dialog.setLayout(layout)
        dialog.exec_()

    def generar_reporte(self):
        """Genera un reporte en formato Excel con los datos filtrados por las fechas seleccionadas."""
        fecha_inicio = self.fecha_inicio_entry.date().toString("yyyy-MM-dd")
        fecha_fin = self.fecha_fin_entry.date().toString("yyyy-MM-dd")

        # Crear la cadena de conexión a la base de datos
        conexion = f"postgresql://{self.USER}:{self.PWD}@{self.SERVER}/{self.BD}?client_encoding=utf8"
        sql = self.sql_query.replace("fecha_inicio", fecha_inicio).replace("fecha_fin", fecha_fin)

        try:
            engine = create_engine(conexion)
            with engine.connect() as connection:
                result = connection.execute(text(sql))
                df = pd.DataFrame(result.fetchall(), columns=result.keys())  # Convertir a DataFrame

            # Guardar el reporte en formato Excel
            nombre_archivo = f"Reporte de Multas {fecha_inicio} a {fecha_fin}.xlsx"
            options = QFileDialog.Options()
            ruta_archivo, _ = QFileDialog.getSaveFileName(
                self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options
            )

            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine="xlsxwriter") as writer:
                    df.to_excel(writer, sheet_name="Concentrado", index=False)

                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el reporte: {e}")

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())


if __name__ == "__main__":
    app = QApplication([])  # Crear aplicación PyQt
    window = ReportMultasWindow()  # Crear ventana del reporte
    window.show()  # Mostrar ventana
    sys.exit(app.exec_())  # Ejecutar la aplicación
