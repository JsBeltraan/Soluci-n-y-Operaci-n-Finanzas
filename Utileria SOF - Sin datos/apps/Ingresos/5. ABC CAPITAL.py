import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QFileDialog, QMessageBox, QDesktopWidget
)

class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        
class ExcelToPostgresInserter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.center_window()

    def initUI(self):
        self.setWindowTitle('Convertir Excel a Insert SQL')
        self.setGeometry(100, 100, 400, 200)


        self.layout = QVBoxLayout(self)

        self.btn_select_folder = QPushButtonBack('Seleccionar Carpeta con Archivos XLS/XLSX', self)
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.layout.addWidget(self.btn_select_folder)

        self.setLayout(self.layout)

    def load_css(self):
        """Carga el archivo CSS para aplicar estilos a los widgets."""
        css_file = "style.css"
        if os.path.exists(css_file):
            with open(css_file, "r") as f:
                self.setStyleSheet(f.read())
        else:
            print("El archivo style.css no se encontró.")

    def select_folder(self):
        """Abrir un diálogo para seleccionar una carpeta y procesar los archivos .xls."""
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
        
        if folder:
            self.process_files_in_folder(folder)

    def process_files_in_folder(self, folder):
        """Leer todos los archivos Excel en la carpeta y generar archivos .txt con inserts."""
        files = [f for f in os.listdir(folder) if f.endswith(('.xls', '.xlsx'))]

        if not files:
            QMessageBox.warning(self, "Sin Archivos", "No se encontraron archivos .xls o .xlsx en la carpeta seleccionada.")
            return

        for file in files:
            file_path = os.path.join(folder, file)
            self.process_file(file_path, folder)

    def process_file(self, file, output_folder):
        """Leer el archivo Excel y generar el archivo .txt con inserts."""
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

            # Abrir el archivo para escribir los inserts
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
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelToPostgresInserter()
    window.show()
    sys.exit(app.exec_())
