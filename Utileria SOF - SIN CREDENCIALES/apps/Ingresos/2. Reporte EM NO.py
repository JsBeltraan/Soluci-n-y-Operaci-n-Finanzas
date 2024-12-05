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

# Clase personalizada para botones
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

class QPushButtonImage(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)

# Ventana principal
class ReportWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()  # Configuración inicial de la interfaz
        self.center_window()  # Centrar la ventana

    def initUI(self):
        self.setWindowTitle('Generador de archivos EM/NO')
        self.setGeometry(100, 100, 800, 350)
        self.setMinimumSize(800, 350)  # Tamaño mínimo de la ventana

        self.load_stylesheet()  # Cargar archivo CSS para estilos

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # Layout para logo y botón de salir
        self.logo_layout = QHBoxLayout()

        # Agregar logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.logo_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para alinear el botón de salir
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.logo_layout.addSpacerItem(spacer)

        # Botón salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)
        self.logo_layout.addWidget(self.btn_salir)

        # Botón para abrir bitácora
        self.btn_bitacora = QPushButtonImage(self)
        icon_pixmap = QPixmap('img/bitacora.png').scaled(20, 20, Qt.KeepAspectRatio)  # Icono de bitácora
        self.btn_bitacora.setIcon(QIcon(icon_pixmap))
        self.btn_bitacora.clicked.connect(self.mostrar_bitacora)
        self.logo_layout.addWidget(self.btn_bitacora)

        # Agregar layout de logo al principal
        self.main_layout.addLayout(self.logo_layout)

        # Layout para botones de empresas
        self.boton_empresa_layout = QVBoxLayout()

        # Botones de cada empresa con su acción asociada
        self.btn_kyara = QPushButtonBack('Kyara', self)
        self.btn_kyara.clicked.connect(lambda: self.generar_archivos("Kyara"))
        self.boton_empresa_layout.addWidget(self.btn_kyara)

        self.btn_ficus = QPushButtonBack('Ficus', self)
        self.btn_ficus.clicked.connect(lambda: self.generar_archivos("Ficus"))
        self.boton_empresa_layout.addWidget(self.btn_ficus)

        self.btn_casas = QPushButtonBack('Casas', self)
        self.btn_casas.clicked.connect(lambda: self.generar_archivos("Casas"))
        self.boton_empresa_layout.addWidget(self.btn_casas)

        # Centrar botones verticalmente con espaciadores
        self.main_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.main_layout.addLayout(self.boton_empresa_layout)
        self.main_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))

        self.setLayout(self.main_layout)

        # Información inicial de la bitácora
        self.server_info = "Servidor: No disponible"
        self.database_info = "Base de datos: No disponible"
        self.queries_info = "No hay consultas ejecutadas aún."

    def load_stylesheet(self):
        """Cargar archivo CSS para estilos."""
        with open('style.css', 'r') as f:
            self.setStyleSheet(f.read())

    def load_logo(self):
        """Cargar logo desde archivo."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: Logo no encontrado en '{logo_path}'.")

    def mostrar_bitacora(self):
        """Mostrar ventana de bitácora."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bitácora")
        dialog.setGeometry(100, 100, 800, 300)
        
        layout = QVBoxLayout()

        # Área de texto para mostrar la bitácora
        bitacora_text = QTextEdit(dialog)
        bitacora_text.setReadOnly(True)
        bitacora_text.setPlainText(
            f"{self.server_info}\n{self.database_info}\n{self.queries_info}"
        )
        layout.addWidget(bitacora_text)

        # Botón para cerrar bitácora
        btn_close = QPushButtonBack("Cerrar", dialog)
        btn_close.clicked.connect(dialog.close)
        layout.addWidget(btn_close)

        dialog.setLayout(layout)

        # Centrar ventana de bitácora
        screen_geometry = QDesktopWidget().availableGeometry()
        dialog_geometry = dialog.geometry()
        x = (screen_geometry.width() - dialog_geometry.width()) // 2
        y = (screen_geometry.height() - dialog_geometry.height()) // 2
        dialog.move(x, y)

        dialog.exec_()

    def generar_archivos(self, Nom_empresa):
        """Generar archivos según la empresa seleccionada."""
        F_EM = f"C:/Archivos_EM_NO/{Nom_empresa}/EM/"
        F_NO = f"C:/Archivos_EM_NO/{Nom_empresa}/NO/"

        # Crear carpetas para guardar archivos
        os.makedirs(F_EM, exist_ok=True)
        os.makedirs(F_NO, exist_ok=True)

        # Configurar parámetros según empresa
        if Nom_empresa == "Kyara":
            server, database, user, pw, port, Num_empresa = "SERVIDOR", "BD", "USUARIO", "PASSWORD", "PUERTO", "EMPRESA"

        elif Nom_empresa == "Ficus":
            server, database, user, pw, port, Num_empresa = "SERVIDOR", "BD", "USUARIO", "PASSWORD", "PUERTO", "EMPRESA"

        elif Nom_empresa == "Casas":
            server, database, user, pw, port, Num_empresa = "SERVIDOR", "BD", "USUARIO", "PASSWORD", "PUERTO", "EMPRESA"

        # Actualizar bitácora con conexión
        self.server_info = f"Servidor: {server}"
        self.database_info = f"Base de datos: {database}"

        # Generar cadena de conexión
        user_encode = urllib.parse.quote_plus(user)
        pw_encode = urllib.parse.quote_plus(pw)
        connection_string = f"postgresql+psycopg2://{user_encode}:{pw_encode}@{server}:{port}/{database}"
        engine = create_engine(connection_string)

        try:
            with engine.connect() as conn:
                # Ejemplo de consultas y creación de archivos
                consulta_nom = "CONSULTA"
                df_nom = pd.read_sql_query(consulta_nom, conn)
                quin = str(df_nom["FECHA"].values[0])

                # Crear archivo EM
                consulta_em = "CONSULTA"
                df_em = pd.read_sql_query(consulta_em, conn)
                df_em.to_csv(f"{F_EM}EM{quin}.{Num_empresa}", index=False, sep="\t", header=False)

                # Crear archivo NO
                consulta_no = "CONSULTA"
                df_no = pd.read_sql_query(consulta_no, conn)
                df_no.to_csv(f"{F_NO}NO{quin}.{Num_empresa}", index=False, sep="\t", header=False)

                # Actualizar bitácora con consultas realizadas
                self.queries_info = (
                    f"Empresa: {Nom_empresa}\n"
                    f"1. {consulta_nom};\n"
                    f"2. {consulta_em};\n"
                    f"3. {consulta_no};\n"
                )

                QMessageBox.information(self, "Éxito", f"Archivos de {Nom_empresa} generados con éxito")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportWindow()
    window.show()
    sys.exit(app.exec_())
