import sys
import os
import subprocess
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QGridLayout, QMessageBox, QSpacerItem, QSizePolicy, QDesktopWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Ventana Principal')
        self.setGeometry(100, 100, 800, 600)
        self.setMinimumSize(800, 600)  # Tamaño mínimo de la ventana
        self.center_window()

        # Cargar el archivo CSS
        self.load_stylesheet()

        # Layout principal vertical
        self.main_layout = QVBoxLayout(self)
        
        # Layout horizontal para el logo y regresar
        self.logo_layout = QHBoxLayout()
        
        # Logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.logo_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)
        
        # Botón de "Regresar" hacia la derecha
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.logo_layout.addSpacerItem(spacer)
        
        # Crear botón "Regresar"
        self.regresar_button = QPushButtonBack("Regresar", self)
        self.regresar_button.clicked.connect(self.mostrar_menu_principal)
        self.logo_layout.addWidget(self.regresar_button, alignment=Qt.AlignRight)
        self.regresar_button.hide()  # Oculto en pestaña inicial
        
        # Logo al layout principal
        self.main_layout.addLayout(self.logo_layout)
        
        # Centrar layout para los botones
        self.button_layout = QVBoxLayout()
        
        # Espaciadores para centrar verticalmente los botones
        self.button_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))  # Espaciador superior

        # Grid para los botones
        self.grid_layout = QGridLayout()
        self.button_layout.addLayout(self.grid_layout)

        # Espaciador inferior para centrar los botones
        self.button_layout.addSpacerItem(QSpacerItem(20, 200, QSizePolicy.Minimum, QSizePolicy.Expanding))

        self.main_layout.addLayout(self.button_layout)
        self.mostrar_menu_principal()

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

    def mostrar_menu_principal(self):
        """Mostrar el menú principal con los botones de las carpetas."""
        self.clear_layout(self.grid_layout)
        self.regresar_button.hide()

        # Directorio de apps
        directorio_principal = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'apps')

        # Validar que la carpeta 'apps' exista
        if not os.path.exists(directorio_principal):
            os.makedirs(directorio_principal)

        # Agregar un botón para cada carpeta en el directorio 'apps' que contenga archivos .py
        carpetas = [d for d in os.listdir(directorio_principal) if os.path.isdir(os.path.join(directorio_principal, d))]

    def mostrar_submenu(self, carpeta):
        """Mostrar el submenú con botones para los programas en la carpeta seleccionada."""
        self.clear_layout(self.grid_layout)

        # Mostrar botón regresar
        self.regresar_button.show()

        # Listar archivos Python en la carpeta
        archivos = os.listdir(carpeta)
        archivos = [a for a in archivos if os.path.isfile(os.path.join(carpeta, a)) and a.endswith('.py')]

        if not archivos:
            QMessageBox.information(self, "Info", "No se encontraron archivos Python en la carpeta.")
            return

    def abrir_programa(self, carpeta, archivo):
        """Abrir un archivo Python específico usando subprocess."""
        ruta = os.path.join(carpeta, archivo)
        subprocess.Popen(["python", ruta])

    def clear_layout(self, layout):
        """Eliminar todos los widgets de un layout."""
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry() # Obtener dimensiones de la ventana
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
