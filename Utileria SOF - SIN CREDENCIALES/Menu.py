import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QGridLayout,
    QMessageBox, QSpacerItem, QSizePolicy, QDesktopWidget
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

# Clase personalizada para botones
class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

# Ventana principal de la aplicación
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Ventana Principal')
        self.setGeometry(100, 100, 800, 600)
        self.setMinimumSize(800, 600)  # Tamaño mínimo de la ventana
        self.center_window()  # Centrar ventana en la pantalla

        # Cargar estilo desde archivo CSS
        self.load_stylesheet()

        # Layout principal vertical
        self.main_layout = QVBoxLayout(self)

        # Layout para el logo y botón "Regresar"
        self.logo_layout = QHBoxLayout()

        # Agregar logo
        self.logo_label = QLabel(self)
        self.load_logo()  # Cargar imagen del logo
        self.logo_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para separar logo del botón "Regresar"
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.logo_layout.addSpacerItem(spacer)

        # Botón "Regresar" (oculto inicialmente)
        self.regresar_button = QPushButtonBack("Regresar", self)
        self.regresar_button.clicked.connect(self.mostrar_menu_principal)
        self.logo_layout.addWidget(self.regresar_button, alignment=Qt.AlignRight)
        self.regresar_button.hide()

        # Añadir layout del logo al principal
        self.main_layout.addLayout(self.logo_layout)

        # Layout central para los botones
        self.button_layout = QVBoxLayout()
        self.button_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))  # Espaciador superior
        self.grid_layout = QGridLayout()  # Botones organizados en cuadrícula
        self.button_layout.addLayout(self.grid_layout)
        self.button_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))  # Espaciador inferior

        # Añadir layout de botones al principal
        self.main_layout.addLayout(self.button_layout)

        # Mostrar el menú principal al iniciar
        self.mostrar_menu_principal()

    def load_stylesheet(self):
        """Carga el archivo de estilo CSS."""
        with open('style.css', 'r') as f:
            style = f.read()
            self.setStyleSheet(style)

    def load_logo(self):
        """Cargar y mostrar el logo."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)  # Ajustar tamaño del logo
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: No se encontró '{logo_path}'.")

    def mostrar_menu_principal(self):
        """Mostrar los botones del menú principal."""
        self.clear_layout(self.grid_layout)  # Limpiar el grid de botones
        self.regresar_button.hide()  # Ocultar botón "Regresar"

        # Directorio donde están las carpetas de las apps
        directorio_principal = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'apps')

        # Crear la carpeta 'apps' si no existe
        if not os.path.exists(directorio_principal):
            os.makedirs(directorio_principal)

        # Buscar carpetas dentro de 'apps'
        carpetas = [d for d in os.listdir(directorio_principal) if os.path.isdir(os.path.join(directorio_principal, d))]

        # Crear un botón por cada carpeta que contenga archivos .py
        num_columnas = 1  # Número de columnas en el grid
        row, col = 0, 0
        for carpeta in carpetas:
            ruta_carpeta = os.path.join(directorio_principal, carpeta)
            archivos_py = [a for a in os.listdir(ruta_carpeta) if a.endswith('.py')]

            if archivos_py:  # Crear botón solo si hay scripts Python
                boton = QPushButton(carpeta, self)
                boton.clicked.connect(lambda checked, c=ruta_carpeta: self.mostrar_submenu(c))
                self.grid_layout.addWidget(boton, row, col)  # Agregar botón al grid
                col += 1
                if col == num_columnas:
                    col = 0
                    row += 1

    def mostrar_submenu(self, carpeta):
        """Mostrar botones para los programas en una carpeta específica."""
        self.clear_layout(self.grid_layout)  # Limpiar el grid
        self.regresar_button.show()  # Mostrar botón "Regresar"

        # Listar archivos Python en la carpeta
        archivos = [a for a in os.listdir(carpeta) if a.endswith('.py')]

        if not archivos:
            QMessageBox.information(self, "Info", "No hay archivos Python en esta carpeta.")
            return

        row, col = 0, 0
        num_columnas = 1
        for archivo in archivos:
            nombre = os.path.splitext(archivo)[0]  # Obtener nombre sin extensión
            boton = QPushButton(nombre, self)
            boton.clicked.connect(lambda checked, f=archivo: self.abrir_programa(carpeta, f))
            self.grid_layout.addWidget(boton, row, col)  # Agregar botón al grid
            col += 1
            if col == num_columnas:
                col = 0
                row += 1

    def abrir_programa(self, carpeta, archivo):
        """Ejecutar un archivo Python con subprocess."""
        ruta = os.path.join(carpeta, archivo)
        subprocess.Popen(["python", ruta])

    def clear_layout(self, layout):
        """Eliminar widgets del layout."""
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
