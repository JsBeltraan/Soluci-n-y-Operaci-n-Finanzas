import sys
import os
import pyodbc
from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout, QMessageBox, QDesktopWidget, QInputDialog
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from Menu import MainWindow

class QPushButtonBack(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)

class LoginSystem(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Login System")
        self.setGeometry(900, 100, 400, 350)
        self.center_window()

        # Conexión a la base de datos
        self.conn = self.connect_db()

        # Cargar archivo de estilo CSS
        self.load_stylesheet()

        # Configuración de la interfaz
        self.layout = QVBoxLayout()

        # Logo en la parte superior izquierda
        self.logo_label = QLabel(self)
        self.load_logo()
        self.layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Campos de entrada
        self.user_label = QLabel("Usuario:")
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Usuario")
        self.layout.addWidget(self.user_label)
        self.layout.addWidget(self.user_input, alignment=Qt.AlignCenter)

        self.pass_label = QLabel("Contraseña:")
        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("Contraseña")
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(self.pass_label)
        self.layout.addWidget(self.pass_input, alignment=Qt.AlignCenter)

        # Botones (mismo tamaño)
        button_width = 200  # Ancho botones
        self.login_button = QPushButtonBack("Iniciar sesión")
        self.login_button.setFixedWidth(button_width)
        self.login_button.clicked.connect(self.login)
        self.layout.addWidget(self.login_button, alignment=Qt.AlignCenter)

        self.register_button = QPushButtonBack("Registrar nuevo usuario")
        self.register_button.setFixedWidth(button_width)
        self.register_button.clicked.connect(self.register_user)
        self.layout.addWidget(self.register_button, alignment=Qt.AlignCenter)

        self.change_pass_button = QPushButtonBack("Cambiar contraseña")
        self.change_pass_button.setFixedWidth(button_width)
        self.change_pass_button.clicked.connect(self.change_password)
        self.layout.addWidget(self.change_pass_button, alignment=Qt.AlignCenter)

        self.setLayout(self.layout)

    def load_stylesheet(self):
        """Carga el archivo de estilo CSS."""
        with open("style.css", "r") as file:
            style = file.read()
            self.setStyleSheet(style)

    def load_logo(self):
        """Cargar el logo."""
        logo_path = "img/Logo-coppel.png"
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def connect_db(self):
        try:
            conn = pyodbc.connect(
                "DRIVER={ODBC Driver 17 for SQL Server};"
                "SERVER=SERVER;"
                "DATABASE=BD;"
                "Trusted_Connection=yes;"
            )
            return conn
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo conectar a la base de datos: {e}")
            sys.exit(1)

    def login(self):
        username = self.user_input.text()
        password = self.pass_input.text()
        
        cursor = self.conn.cursor()
        cursor.execute("SELECT role FROM users WHERE username = ? AND password = ?", (username, password))
        result = cursor.fetchone()
        
        if result:
            self.user_role = result[0]  # Se guarda rol
            self.open_menu()  # Abrir ventana principal y se cierra login.py
        else:
            QMessageBox.warning(self, "Login Fallido", "Usuario o contraseña incorrectos")

    def open_menu(self):
        """Cerrar la ventana de inicio de sesión y abrir el menú."""
        self.close()
        self.menu_window = MainWindow()
        self.menu_window.show()

    def register_user(self):
        if getattr(self, 'user_role', None) != 'admin':
            QMessageBox.warning(self, "Acceso denegado", "Solo un administrador puede registrar nuevos usuarios.")
            return

        new_username, ok = QInputDialog.getText(self, "Registrar usuario", "Ingresa el nombre de usuario:")
        if not ok or not new_username:
            return

        new_password, ok = QInputDialog.getText(self, "Registrar usuario", "Ingresa la contraseña:", QLineEdit.Password)
        if not ok or not new_password:
            return

        cursor = self.conn.cursor()
        try:
            cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (new_username, new_password, 'user'))
            self.conn.commit()
            QMessageBox.information(self, "Registro exitoso", f"Usuario '{new_username}' registrado correctamente.")
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"No se pudo registrar el usuario: {e}")

    def change_password(self):
        username = self.user_input.text()
        current_password = self.pass_input.text()

        if not username or not current_password:
            QMessageBox.warning(self, "Error", "Debe ingresar el usuario y la contraseña actuales.")
            return

        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, current_password))
        result = cursor.fetchone()

        if not result:
            QMessageBox.warning(self, "Error", "Usuario o contraseña incorrectos.")
            return

        new_password, ok = QInputDialog.getText(self, "Cambiar contraseña", "Ingresa la nueva contraseña:", QLineEdit.Password)
        if not ok or not new_password:
            return

        try:
            cursor.execute("UPDATE users SET password = ? WHERE username = ?", (new_password, username))
            self.conn.commit()
            QMessageBox.information(self, "Cambio exitoso", "Contraseña actualizada correctamente.")
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar la contraseña: {e}")

    def center_window(self):
        """Centrar la ventana en la pantalla."""
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Estilos
    with open("style.css", "r") as file:
        app.setStyleSheet(file.read())

    window = LoginSystem()
    window.show()
    sys.exit(app.exec_())
