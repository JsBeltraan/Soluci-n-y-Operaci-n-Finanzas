import os
from datetime import datetime, timedelta
import pandas as pd
from sqlalchemy import create_engine
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QDateEdit, QHBoxLayout, QSpacerItem, QSizePolicy, QComboBox, QCheckBox, QTextEdit, QDialog
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QDate
import sys

# Clase personalizada para un botón
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

# Clase personalizada para un botón con una imagen
class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

# Clase personalizada para el ComboBox
class QComboBoxNew(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)

# Ventana principal para seleccionar reportes
class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Seleccionar Reporte")  # Título de la ventana
        self.setMinimumSize(600, 500)  # Tamaño mínimo de la ventana
        self.center_window()  # Centra la ventana en la pantalla

        # Información de conexión a la base de datos
        self.SERVER = "SERVIDOR"
        self.BD = "BASEDEDATOS"
        self.USER = "USUARIO"
        self.PWD = "CONTRASEÑA"
        self.PORT = "PUERTO"

        # Variable para almacenar la consulta SQL
        self.sql_query = ""

        # Establecer conexión con la base de datos PostgreSQL usando SQLAlchemy
        self.engine = create_engine(f'postgresql://{self.USER}:{self.PWD}@{self.SERVER}:{self.PORT}/{self.BD}?client_encoding=utf-8')

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # Layout para el header (logo y botón de bitácora)
        self.header_layout = QHBoxLayout()
        self.logo_label = QLabel(self)
        self.load_logo()  # Cargar el logo
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para empujar el botón de bitácora hacia la derecha
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)
        self.header_layout.addWidget(self.btn_salir)

        # Botón para mostrar la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.header_layout.addWidget(self.btn_bitacora)

        # Añadir el layout del header al layout principal
        self.main_layout.addLayout(self.header_layout)

        # Selección de tipo de reporte
        self.reporte_label = QLabel("Seleccione el Tipo de Reporte:", self)
        self.main_layout.addWidget(self.reporte_label)

        self.tipo_reporte = QComboBoxNew(self)
        self.tipo_reporte.addItems(["Transferencias", "Corresponsales"])  # Opciones para el tipo de reporte
        self.main_layout.addWidget(self.tipo_reporte)

        # Selección de periodo
        self.periodo_label = QLabel("Seleccione el Periodo:", self)
        self.main_layout.addWidget(self.periodo_label)

        self.opcion_periodo = QComboBoxNew(self)
        self.opcion_periodo.addItems(["Ultimo mes", "Por rango de fechas"])  # Opciones para el periodo
        self.main_layout.addWidget(self.opcion_periodo)

        # Checkbox para usar rango de fechas
        self.rango_fecha_checkbox = QCheckBox("¿Usar rango de fechas?", self)
        self.rango_fecha_checkbox.stateChanged.connect(self.toggle_fecha_inputs)  # Conectar el evento para mostrar/ocultar campos de fechas
        self.main_layout.addWidget(self.rango_fecha_checkbox)

        # Layout horizontal para fechas de inicio y fin
        self.fecha_layout = QHBoxLayout()

        # Fecha de inicio
        self.fecha_inicio_label = QLabel("Fecha inicio:", self)
        self.fecha_layout.addWidget(self.fecha_inicio_label)

        self.fecha_inicio_input = QDateEdit(self)
        self.fecha_inicio_input.setCalendarPopup(True)
        self.fecha_inicio_input.setDate(QDate.currentDate().addMonths(-1))  # Valor por defecto
        self.fecha_layout.addWidget(self.fecha_inicio_input)

        # Fecha fin
        self.fecha_fin_label = QLabel("Fecha fin:", self)
        self.fecha_layout.addWidget(self.fecha_fin_label)

        self.fecha_fin_input = QDateEdit(self)
        self.fecha_fin_input.setCalendarPopup(True)
        self.fecha_fin_input.setDate(QDate.currentDate())  # Valor por defecto
        self.fecha_layout.addWidget(self.fecha_fin_input)

        self.main_layout.addLayout(self.fecha_layout)

        # Inicialmente ocultar los campos de fecha hasta que se marque el checkbox
        self.toggle_fecha_inputs(self.rango_fecha_checkbox.checkState())  

        # Botón para generar el reporte
        self.generar_reporte_button = QPushButtonBack("Generar Reporte", self)
        self.generar_reporte_button.clicked.connect(self.generar_reporte)
        self.main_layout.addWidget(self.generar_reporte_button, alignment=Qt.AlignCenter)

        # Cargar estilos desde un archivo CSS
        self.load_css()

    # Función para mostrar u ocultar los campos de fecha según el checkbox
    def toggle_fecha_inputs(self, state):
        is_checked = (state == Qt.Checked)
        self.fecha_inicio_label.setVisible(is_checked)
        self.fecha_inicio_input.setVisible(is_checked)
        self.fecha_fin_label.setVisible(is_checked)
        self.fecha_fin_input.setVisible(is_checked)

    # Cargar el archivo CSS para aplicar estilos a los widgets
    def load_css(self):
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())
        else:
            print("El archivo style.css no se encontró.")

    # Cargar el logo desde un archivo de imagen
    def load_logo(self):
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    # Mostrar la bitácora con información sobre la conexión
    def mostrar_bitacora(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setMinimumSize(600, 400)

        # Layout para mostrar la bitácora
        layout = QVBoxLayout()

        # Widget QTextEdit para mostrar la bitácora
        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)
        bitacora_text.setPlainText(
            f"Servidor: {self.SERVER}\nBase de Datos: {self.BD}\nConsulta ejecutada:\n{self.sql_query}"
        )
        layout.addWidget(bitacora_text)

        dialog.setLayout(layout)
        dialog.exec_()

    # Función para generar el reporte según la configuración seleccionada
    def generar_reporte(self):
        tipo = self.tipo_reporte.currentText()  # Tipo de reporte seleccionado
        id_producto = '112' if tipo == 'Transferencias' else '056'  # ID según tipo de reporte
        fecha_inicio = self.fecha_inicio_input.date().toString("yyyy-MM-dd") if self.rango_fecha_checkbox.isChecked() else None
        fecha_fin = self.fecha_fin_input.date().toString("yyyy-MM-dd") if self.rango_fecha_checkbox.isChecked() else None
        por_mes = self.opcion_periodo.currentText() == 'Por mes'

        # Consulta SQL base
        self.sql_query = f"CONSULTA 1"

        # Modificar la consulta según el periodo seleccionado
        if por_mes:
            self.sql_query += f"CONSULTA 2"
        elif fecha_inicio and fecha_fin:
            self.sql_query += f"CONSULTA 3"
        
        self.sql_query += "CONSULTA 4"

        # Ejecutar la consulta y guardar los resultados en un archivo Excel
        try:
            with self.engine.connect() as connection:
                df = pd.read_sql_query(self.sql_query, connection)
                nombre_archivo = f'Reporte_{tipo}.xlsx'

                # Mostrar un cuadro de diálogo para guardar el archivo
                options = QFileDialog.Options()
                ruta_archivo, _ = QFileDialog.getSaveFileName(self, "Guardar reporte", nombre_archivo, "Excel Files (*.xlsx)", options=options)

                if ruta_archivo:
                    # Escribir el reporte en un archivo Excel
                    with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                        df.to_excel(writer, sheet_name=tipo, index=False)

                        workbook = writer.book
                        worksheet = writer.sheets[tipo]

                        # Ajustar el ancho de las columnas
                        for i, col in enumerate(df.columns):
                            max_len = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
                            worksheet.set_column(i, i, max_len)

                    # Mostrar mensaje de éxito
                    QMessageBox.information(self, "Reporte Generado", "El reporte ha sido generado exitosamente.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Hubo un problema al generar el reporte: {str(e)}")

    # Centrar la ventana en la pantalla
    def center_window(self):
        screen = QApplication.primaryScreen()
        rect = screen.availableGeometry()
        self.move((rect.width() - self.width()) // 2, (rect.height() - self.height()) // 2)

# Función principal para iniciar la aplicación
def main():
    app = QApplication(sys.argv)
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
