import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout,
    QFileDialog, QMessageBox, QDesktopWidget, QHBoxLayout, QSpacerItem, QSizePolicy
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

class QPushButtonBack(QPushButton):
    """Botón personalizado para usar en la aplicación."""
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

class ExcelToPostgresInserter(QWidget):
    """Ventana principal para convertir archivos Excel a sentencias SQL para PostgreSQL."""
    def __init__(self):
        super().__init__()
        self.initUI()  # Inicializa la interfaz de usuario
        self.center_window()  # Centra la ventana en la pantalla

    def initUI(self):
        """Configura los widgets y el layout de la interfaz gráfica."""
        self.setWindowTitle('Convertir Excel a Insert SQL')
        self.setMinimumSize(600, 400)

        # Layout principal
        self.layout = QVBoxLayout(self)

        # Layout para el header (logo y botón salir)
        self.header_layout = QHBoxLayout()

        # Label para el logo
        self.logo_label = QLabel(self)
        self.load_logo()  # Carga el logo
        self.header_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Espaciador para empujar el botón "Salir" a la derecha
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addSpacerItem(spacer)

        # Botón de salir
        self.btn_salir = QPushButtonBack('Salir', self)
        self.btn_salir.clicked.connect(self.close)  # Conecta la acción de cerrar la ventana
        self.header_layout.addWidget(self.btn_salir)

        # Añadir el layout del header al layout principal
        self.layout.addLayout(self.header_layout)

        # Layout para el contenido principal
        self.content_layout = QVBoxLayout()

        # Espaciador flexible para alinear el contenido centrado verticalmente
        self.content_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Botón de seleccionar carpeta
        self.btn_select_folder = QPushButtonBack('Seleccionar Carpeta con Archivos XLS/XLSX', self)
        self.btn_select_folder.clicked.connect(self.select_folder)  # Abre el selector de carpeta
        self.content_layout.addWidget(self.btn_select_folder, alignment=Qt.AlignCenter)

        # Espaciador flexible para alinear el contenido centrado verticalmente
        self.content_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Añadir el layout del contenido al layout principal
        self.layout.addLayout(self.content_layout)

        # Aplicar estilos desde archivo CSS
        self.load_css()

    def load_css(self):
        """Carga el archivo CSS para aplicar estilos a los widgets."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())  # Aplica los estilos CSS
        else:
            print("El archivo style.css no se encontró.")

    def load_logo(self):
        """Cargar el logo en la interfaz."""
        logo_path = 'img/Logo-coppel.png'
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def select_folder(self):
        """Abrir un diálogo para seleccionar una carpeta y procesar los archivos .xls."""
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
        
        if folder:
            self.process_files_in_folder(folder)  # Procesa los archivos de la carpeta seleccionada

    def process_files_in_folder(self, folder):
        """Leer todos los archivos Excel en la carpeta y generar archivos .txt con inserts."""
        files = [f for f in os.listdir(folder) if f.endswith(('.xls', '.xlsx'))]

        if not files:
            QMessageBox.warning(self, "Sin Archivos", "No se encontraron archivos .xls o .xlsx en la carpeta seleccionada.")
            return

        for file in files:
            file_path = os.path.join(folder, file)
            self.process_file(file_path, folder)  # Procesa cada archivo Excel

    def process_file(self, file, output_folder):
        """Leer el archivo Excel y generar el archivo .txt con inserts SQL."""
        try:
            # Leer el archivo Excel
            df = pd.read_excel(file)

            # Omitir la primera columna
            df = df.iloc[:, 1:]

            # Asegurarse de que solo se tome la fecha sin la hora en el campo de fecha
            for col in df.select_dtypes(include=['datetime']):
                df[col] = df[col].dt.date  # Solo conservar la parte de la fecha

            # Obtener el nombre del archivo sin extensión
            base_name = os.path.splitext(os.path.basename(file))[0]

            # Crear el nombre del archivo .txt
            txt_file_name = f"{base_name}.txt"
            txt_file_path = os.path.join(output_folder, txt_file_name)

            # Abrir el archivo para escribir los inserts SQL
            with open(txt_file_path, 'w') as f:
                for index, row in df.iterrows():
                    # Generar la sentencia INSERT para cada fila
                    columns = ', '.join(row.index)
                    values = ', '.join([f"'{str(v)}'" for v in row.values])
                    insert_statement = f"INSERT INTO TABLA VALUES ({values});\n"
                    f.write(insert_statement)

            QMessageBox.information(self, "Éxito", f"Archivo {txt_file_name} generado con éxito en la carpeta seleccionada.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al procesar {file}: {str(e)}")

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()  # Obtener dimensiones de la ventana
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())  # Mover la ventana al centro de la pantalla

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelToPostgresInserter()  # Crear la ventana principal
    window.show()  # Mostrar la ventana
    sys.exit(app.exec_())  # Ejecutar la aplicación PyQt
