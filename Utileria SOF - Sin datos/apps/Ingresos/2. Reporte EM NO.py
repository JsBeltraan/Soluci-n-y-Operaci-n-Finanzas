import sys
import os
import pandas as pd
import urllib.parse
from sqlalchemy import create_engine
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout,
    QHBoxLayout, QMessageBox, QSpacerItem, QSizePolicy, QDialog,
    QTextEdit, QDesktopWidget
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QIcon


class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)


class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)


class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.center_window()

    def initUI(self):
        self.setWindowTitle('Generador de archivos EM/NO')
        self.setGeometry(100, 100, 800, 350)
        self.setMinimumSize(800, 350)

        # Cargar el archivo CSS
        self.load_stylesheet()

        # Layout principal vertical
        self.main_layout = QVBoxLayout(self)

        # Layout horizontal para el logo y el botón regresar
        self.logo_layout = QHBoxLayout()

        # Label para el logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.logo_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para empujar el botón de "Regresar" hacia la derecha
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.logo_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)
        self.logo_layout.addWidget(self.btn_salir)

        # Botón para abrir la bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)  # Ruta icono de la libreta
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.logo_layout.addWidget(self.btn_bitacora)

        # Añadir el layout de logo al layout principal
        self.main_layout.addLayout(self.logo_layout)

        # Layout para los botones de empresas
        self.boton_empresa_layout = QVBoxLayout()

        # Botones para cada empresa
        self.btn_kyara = QPushButtonBack('Kyara', self)
        self.btn_kyara.clicked.connect(lambda: self.generar_archivos("Kyara"))
        self.boton_empresa_layout.addWidget(self.btn_kyara)

        self.btn_ficus = QPushButtonBack('Ficus', self)
        self.btn_ficus.clicked.connect(lambda: self.generar_archivos("Ficus"))
        self.boton_empresa_layout.addWidget(self.btn_ficus)

        self.btn_casas = QPushButtonBack('Casas', self)
        self.btn_casas.clicked.connect(lambda: self.generar_archivos("Casas"))
        self.boton_empresa_layout.addWidget(self.btn_casas)

        # Centrar verticalmente los botones
        self.main_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))  # Espaciador superior
        self.main_layout.addLayout(self.boton_empresa_layout)
        self.main_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))  # Espaciador inferior

        self.setLayout(self.main_layout)

        # Información sobre servidor, base de datos y consultas - Antes de ejecutar consultas
        self.server_info = "Servidor: No disponible" 
        self.database_info = "Base de datos: No disponible"  
        self.queries_info = "No hay consultas ejecutadas aún."

    def load_stylesheet(self):
        """Carga el archivo de estilo CSS."""
        with open('style.css', 'r') as f:
            style = f.read()
            self.setStyleSheet(style)

    def load_logo(self):
        """Cargar el logo."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def mostrar_bitacora(self):
        """Mostrar el cuadro de diálogo de la bitácora."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setGeometry(100, 100, 400, 300)

        layout = QVBoxLayout()

        # Texto con la información que se mostrara
        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)
        bitacora_text.setPlainText(
            f"{self.server_info}\n"
            f"{self.database_info}\n"
            "Consultas ejecutadas:\n"
            f"{self.queries_info}"
        )
        layout.addWidget(bitacora_text)

        # Botón para cerrar el diálogo
        btn_close = QPushButtonBack("Cerrar", dialog)
        btn_close.clicked.connect(dialog.close)
        layout.addWidget(btn_close)

        dialog.setLayout(layout)
        dialog.exec_()

    def generar_archivos(self, Nom_empresa):
        """Generar archivos EM/NO según la empresa seleccionada."""
        F_EM = f"C:/Archivos_EM_NO/{Nom_empresa}/EM/"
        F_NO = f"C:/Archivos_EM_NO/{Nom_empresa}/NO/"

        try:
            os.makedirs(F_EM, exist_ok=True)
            os.makedirs(F_NO, exist_ok=True)
        except OSError as e:
            print(f"Error al crear carpetas: {e}")

        # Definir los parámetros de conexión según la empresa seleccionada
        if Nom_empresa == "EMPRESA":
            self.server_info = "Servidor: 00.00.00.00"
            self.database_info = "Base de datos: BD"
            user="USER"
            pw="PASS"
            port="0000"
            Num_empresa = "000"

        elif Nom_empresa == "EMPRESA":
            self.server_info = "Servidor: 00.00.00.00"
            self.database_info = "Base de datos: BD"
            user="USER"
            pw="PASS"
            port="0000"
            Num_empresa = "000"

        elif Nom_empresa == "EMPRESA":
            self.server_info = "Servidor: 00.00.00.00"
            self.database_info = "Base de datos: BD"
            user="USER"
            pw="PASS"
            port="0000"
            Num_empresa = "000"

        user_encode=urllib.parse.quote_plus(user)
        pw_encode=urllib.parse.quote_plus(pw)

        # Crear la cadena de conexión
        connection_string = f"postgresql+psycopg2://{user_encode}:{pw_encode}@{self.server_info.split(': ')[1]}:{port}/{self.database_info.split(': ')[1]}"
        engine = create_engine(connection_string)
        print(engine)

        # Actualizar la bitácora con los detalles de la conexión
        self.queries_info = (
            f"Empresa: {Nom_empresa}\n"
            f"{self.server_info}\n"
            f"{self.database_info}\n"
            "Consultas ejecutadas:\n"
            "1. CONSULTA SQL PARA EM;\n"
            "2. CONSULTA SQL PARA NO;\n"
        )

        try:
            with engine.connect() as conn:
                QMessageBox.information(self, "Conexión Exitosa", f"Conectado a {Nom_empresa} exitosamente")
        except Exception as e:
            QMessageBox.critical(self, "Error de Conexión", str(e))

        #  Consultas y generación de archivos:
        with engine.connect() as conn:
            df_nom = pd.read_sql_query("CONSULTA", conn)
            quin = str(df_nom["fechanomina"].values[0])

            # Generar archivo EM
            sql_em = "CONSULTA"
            df_em = pd.read_sql_query(sql_em, conn)
            df_em.to_csv(f"{F_EM}EM{quin}.{Num_empresa}", index=False, sep="\t", header=False)

            # Generar archivo NO
            sql_no = "CONSULTA"
            df_no = pd.read_sql_query(sql_no, conn)
            df_no.to_csv(f"{F_NO}NO{quin}.{Num_empresa}", index=False, sep="\t", header=False)

        QMessageBox.information(self, "Generador de archivos EM/NO", f"Archivos de {Nom_empresa} generados con éxito")

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()  # Obtener dimensiones de la ventana
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())
