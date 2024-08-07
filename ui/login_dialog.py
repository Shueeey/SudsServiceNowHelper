from PyQt5.QtWidgets import QDialog, QLineEdit, QPushButton, QLabel, QVBoxLayout, QMessageBox

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Login")
        self.username = QLineEdit(self)
        self.password = QLineEdit(self)
        self.password.setEchoMode(QLineEdit.Password)
        self.login_button = QPushButton('Login', self)
        self.login_button.clicked.connect(self.handle_login)
        self.username_label = QLabel('Username:')
        self.password_label = QLabel('Password:')

        layout = QVBoxLayout()
        layout.addWidget(self.username_label)
        layout.addWidget(self.username)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password)
        layout.addWidget(self.login_button)
        self.setLayout(layout)

        self.setStyleSheet("""
            QDialog {
                background-color: #1c2432;
                color: #ffffff;
            }
            QLabel {
                color: #a7b3c6;
                font-size: 14px;
            }
            QLineEdit {
                background-color: #2f3c4e;
                border: 1px solid #4a5d78;
                padding: 5px;
                border-radius: 3px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 3px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)

    def handle_login(self):
        if self.username.text() and self.password.text():
            self.accept()
        else:
            QMessageBox.warning(self, 'Error', 'Both fields are required')

    def get_credentials(self):
        return self.username.text(), self.password.text()