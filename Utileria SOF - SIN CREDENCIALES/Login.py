import sys
import os
import pyodbc
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout, QMessageBox, QDesktopWidget, QInputDialog, QDialog
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from Menu import MainWindow  # Importa la clase MainWindow desde menu.py


class QPushButtonBack(QPushButton):  # Botón personalizado
    def __init__(self, text, parent=None):
        super().__init__(text, parent)


class NewUserWindow(QDialog):  # Ventana para registrar nuevos usuarios
    def __init__(self, conn):
        super().__init__()
        self.conn = conn
        self.setWindowTitle("Registrar Nuevo Usuario")  # Título de la ventana
        self.setGeometry(950, 150, 300, 200)  # Posición y tamaño de la ventana
        layout = QVBoxLayout()

        # Campo para ingresar el nombre de usuario
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Nuevo Usuario")  # Texto
        layout.addWidget(self.username_input)

        # Campo para ingresar la contraseña
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Nueva Contraseña")
        self.password_input.setEchoMode(QLineEdit.Password)  # Ocultar contraseña
        layout.addWidget(self.password_input)

        # Botón para registrar al usuario
        self.submit_button = QPushButtonBack("Registrar Usuario")
        self.submit_button.clicked.connect(self.add_user)  # Conectar botón a función
        layout.addWidget(self.submit_button)

        self.setLayout(layout)

    def add_user(self):  # Función para registrar un nuevo usuario en la base de datos
        new_username = self.username_input.text()
        new_password = self.password_input.text()

        if not new_username or not new_password:  # Validar campos vacíos
            QMessageBox.warning(self, "Error", "Debe completar todos los campos.")
            return

        cursor = self.conn.cursor()
        try:
            # Insertar usuario en la base de datos
            cursor.execute("INSERT INTO TABLA (DATO1, DATO2, DATO3) VALUES (?, ?, ?)", (new_username, new_password, 'user'))
            self.conn.commit()
            QMessageBox.information(self, "Registro exitoso", f"Usuario '{new_username}' registrado correctamente.")
            self.accept()  # Cierra la ventana
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"No se pudo registrar el usuario: {e}")


class LoginSystem(QWidget):  # Ventana principal del sistema de inicio de sesión
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Login System")  # Título de la ventana
        self.setGeometry(900, 100, 400, 350)  # Posición y tamaño
        self.center_window()  # Centrar ventana

        # Conexión a la base de datos
        self.conn = self.connect_db()

        # Cargar archivo CSS para el diseño
        self.load_stylesheet()

        # Configuración de la interfaz
        self.layout = QVBoxLayout()

        # Cargar y agregar el logo
        self.logo_label = QLabel(self)
        self.load_logo()
        self.layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Etiquetas y campos de texto para usuario
        self.user_label = QLabel("Usuario:")
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Usuario")  # Texto de ayuda
        self.layout.addWidget(self.user_label)
        self.layout.addWidget(self.user_input, alignment=Qt.AlignCenter)

        # Etiquetas y campos de texto para contraseña
        self.pass_label = QLabel("Contraseña:")
        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("Contraseña")
        self.pass_input.setEchoMode(QLineEdit.Password)  # Ocultar texto
        self.layout.addWidget(self.pass_label)
        self.layout.addWidget(self.pass_input, alignment=Qt.AlignCenter)

        # Botones con ancho fijo
        button_width = 200
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

    def load_stylesheet(self):  # Cargar diseño CSS
        with open("style.css", "r") as file:
            style = file.read()
            self.setStyleSheet(style)

    def load_logo(self):  # Cargar y configurar logo
        logo_path = "img/Logo-coppel.png"
        if os.path.isfile(logo_path):
            pixmap = QPixmap(logo_path).scaled(145, 30, Qt.KeepAspectRatio)
            self.logo_label.setPixmap(pixmap)
        else:
            print(f"Error: El archivo de imagen '{logo_path}' no se encuentra.")

    def connect_db(self):  # Conectar a la base de datos
        try:
            conn = pyodbc.connect(
                "DRIVER={ODBC Driver 17 for SQL Server};"
                "SERVER=SERVIDOR;"  # Servidor
                "DATABASE=BASEDATOS;"  # Base de datos
                "Trusted_Connection=yes;" # Conexion con usuario de dominio
            )
            return conn
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo conectar a la base de datos: {e}")
            sys.exit(1)

    def login(self):  # Función para iniciar sesión
        username = self.user_input.text()
        password = self.pass_input.text()

        cursor = self.conn.cursor()
        cursor.execute("SELECT role FROM users WHERE username = ? AND password = ?", (username, password))
        result = cursor.fetchone()

        if result:
            self.user_role = result[0]  # Guardar rol de usuario
            self.open_menu()  # Abrir menú principal
        else:
            QMessageBox.warning(self, "Login Fallido", "Usuario o contraseña incorrectos")

    def open_menu(self):  # Abrir ventana del menú
        self.close()
        self.menu_window = MainWindow()
        self.menu_window.show()

    def register_user(self):  # Registrar un nuevo usuario
        # Verificar credenciales de administrador
        admin_user, ok_user = QInputDialog.getText(self, "Autenticación requerida", "Usuario de administrador:")
        if not ok_user or not admin_user:
            return

        admin_pass, ok_pass = QInputDialog.getText(self, "Autenticación requerida", "Contraseña de administrador:", QLineEdit.Password)
        if not ok_pass or not admin_pass:
            return

        cursor = self.conn.cursor()
        cursor.execute("SELECT role FROM TABLA WHERE DATO1 = ? AND DATO2 = ?", (admin_user, admin_pass))
        admin_data = cursor.fetchone()

        if not admin_data or admin_data[0] != 'admin':
            QMessageBox.warning(self, "Acceso denegado", "Credenciales de administrador incorrectas.")
            return

        # Abrir ventana para registrar usuario
        new_user_window = NewUserWindow(self.conn)
        new_user_window.exec_()

    def change_password(self):  # Cambiar contraseña del usuario
        username = self.user_input.text()
        current_password = self.pass_input.text()

        if not username or not current_password:
            QMessageBox.warning(self, "Error", "Debe ingresar el usuario y la contraseña actuales.")
            return

        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM TABLA WHERE DATO1 = ? AND DATO2 = ?", (username, current_password))
        result = cursor.fetchone()

        if not result:
            QMessageBox.warning(self, "Error", "Usuario o contraseña incorrectos.")
            return

        new_password, ok = QInputDialog.getText(self, "Cambiar contraseña", "Ingresa la nueva contraseña:", QLineEdit.Password)
        if not ok or not new_password:
            return

        try:
            cursor.execute("UPDATE TABLA SET DATO1 = ? WHERE DATO2 = ?", (new_password, username))
            self.conn.commit()
            QMessageBox.information(self, "Cambio exitoso", "Contraseña actualizada correctamente.")
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar la contraseña: {e}")

    def center_window(self):  # Centrar la ventana en la pantalla
        qr = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry(self).center()
        qr.moveCenter(screen_center)
        self.move(qr.topLeft())


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Cargar y aplicar archivo CSS
    with open("style.css", "r") as file:
        app.setStyleSheet(file.read())

    window = LoginSystem()
    window.show()
    sys.exit(app.exec_())
