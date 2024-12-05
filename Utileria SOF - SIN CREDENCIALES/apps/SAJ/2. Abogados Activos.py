import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine, text
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QDesktopWidget, QDialog, QTextEdit, QSizePolicy
)
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt
import sys
import traceback

# Clase para botón
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFixedSize(160, 40)  # Ajuste del tamaño para que el texto quepa completo
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Ajusta el tamaño automáticamente

# Clase para el botón con una imagen
class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(40, 40)  # Tamaño fijo pequeño para el botón de la bitácora
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Ajusta el tamaño automáticamente

# Ventana principal que contiene los widgets y funcionalidad
class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reporte de Abogados Activos")
        self.setFixedSize(400, 300)  # Establece tamaño fijo para la ventana
        self.center_window()  # Centra la ventana en la pantalla

        # Atributos de conexión y consulta SQL
        self.engine = None
        self.sql_query = """       
        CONSULTA
        """

        # Layout principal de la ventana
        self.main_layout = QVBoxLayout(self)

        # Layout horizontal para el logo y los botones
        self.header_layout = QHBoxLayout()

        # Cargar logo en la ventana
        self.logo_label = QLabel(self)
        self.load_logo()
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)
        self.header_layout.addWidget(self.btn_salir)

        # Botón para abrir la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)  # Escala la imagen
        icon = QIcon(icon_pixmap)  # Convierte el pixmap a icono
        self.btn_bitacora.setIcon(icon)
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.header_layout.addWidget(self.btn_bitacora)

        # Añadir el layout de cabecera al layout principal
        self.main_layout.addLayout(self.header_layout)

        # Botón para generar reporte
        self.generar_reporte_button = QPushButtonBack("Generar reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Cargar y aplicar estilos desde archivo CSS
        self.load_css()

    # Cargar los estilos CSS desde un archivo
    def load_css(self):
        """Carga el archivo CSS para aplicar estilos a los widgets."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())
        else:
            print("El archivo style.css no se encontró.")

    # Cargar el logo en la interfaz
    def load_logo(self):
        """Cargar el logo en la interfaz."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)  # Escala el logo
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    # Mostrar la bitácora en un cuadro de diálogo
    def mostrar_bitacora(self):
        """Mostrar el cuadro de diálogo de la bitácora centrado en la pantalla."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)

        # Obtener la resolución de la pantalla y centrar el cuadro de diálogo
        screen_geometry = QApplication.primaryScreen().geometry()
        dialog.show()

        layout = QVBoxLayout()

        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)  # Solo lectura

        # Información sobre la conexión a la base de datos
        if self.engine:
            connection_info = f"Servidor: {self.engine.url.host}\nBase de datos: {self.engine.url.database}"
        else:
            connection_info = "Conexión no establecida."

        # Mostrar la información relevante en la bitácora
        bitacora_text.setPlainText(
            f"Información de conexión:\n\n"
            f"{connection_info}\n\n"
            f"Consulta ejecutada:\n"
            f"{self.sql_query}"
        )
        layout.addWidget(bitacora_text)

        dialog.setLayout(layout)
        dialog.exec_()

    # Función para generar el reporte en Excel
    def generar_reporte(self):
        """Generar el reporte en formato Excel a partir de la consulta SQL."""
        print("Inicia reporte")
        sql = self.sql_query  # Consulta SQL definida
        self.engine = create_engine(
            'postgresql://USUARIO:PASSWORD@SERVIDOR:PUERTO/BASEDEDATOS?client_encoding=utf8'  # Cadena de conexión
        )
        try:
            with self.engine.connect() as connection:
                result = connection.execute(text(sql))  # Ejecutar consulta
                df = pd.DataFrame(result.fetchall(), columns=result.keys())  # Convertir resultado en DataFrame

            # Obtener ruta para guardar el archivo
            nombre_archivo = 'Reporte de Abogados Activos.xlsx'
            ruta_archivo, _ = QFileDialog.getSaveFileName(
                self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)"
            )

            if ruta_archivo:
                with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Concentrado')  # Guardar el DataFrame en Excel
                QMessageBox.information(self, "Éxito", "El reporte fue generado con éxito.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar reporte: {e}")
            print(traceback.format_exc())  # Imprimir el error detallado
        finally:
            if self.engine:
                self.engine.dispose()  # Cerrar la conexión

    # Centrar la ventana en la pantalla
    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

# Código principal para ejecutar la aplicación
if __name__ == "__main__":
    app = QApplication([])  # Iniciar la aplicación Qt
    window = ReportWindow()  # Crear ventana principal
    window.show()  # Mostrar la ventana
    sys.exit(app.exec_())  # Ejecutar la aplicación
