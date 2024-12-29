import os
import win32com.client
from subprocess import call
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QFileDialog, QVBoxLayout, QWidget, QMessageBox,
    QRadioButton, QHBoxLayout
)
from PyQt5.QtGui import QPalette, QColor, QIcon
from PyQt5.QtCore import Qt
from pythoncom import IID_IPersistFile, CoCreateInstance, CLSCTX_INPROC_SERVER
from win32com.shell import shell
import ctypes


class HubarLockApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.path = ''
        self.password = ''
        self.is_folder = True

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Blocking directories | files")
        self.setGeometry(200, 200, 400, 300)
        self.setWindowIcon(QIcon('icon.png'))
        self.setAcceptDrops(True)

        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #2b2b2b;
            }
            QLabel {
                color: #f0f0f0;
            }
            QLineEdit {
                background-color: #3c3f41;
                color: #f0f0f0;
                border: 1px solid #555;
                padding: 5px;
            }
            QPushButton {
                background-color: #555;
                color: #f0f0f0;
                border: 1px solid #777;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #777;
            }
            QMessageBox {
                background-color: #3c3f41;
                color: #f0f0f0;
            }
            QRadioButton {
                color: #f0f0f0;
            }
            """
        )

        self.layout = QVBoxLayout()

        self.radio_layout = QHBoxLayout()
        self.folder_radio = QRadioButton('Folder')
        self.folder_radio.setChecked(True)
        self.folder_radio.toggled.connect(self.update_selection_type)
        self.file_radio = QRadioButton('File')
        self.file_radio.toggled.connect(self.update_selection_type)

        self.radio_layout.addWidget(self.folder_radio)
        self.radio_layout.addWidget(self.file_radio)
        self.layout.addLayout(self.radio_layout)

        self.path_label = QLabel("Select folder/file or drag here:")
        self.layout.addWidget(self.path_label)

        self.path_input = QLineEdit()
        self.layout.addWidget(self.path_input)

        self.browse_button = QPushButton('Review')
        self.browse_button.clicked.connect(self.choose_path)
        self.layout.addWidget(self.browse_button)

        self.password_label = QLabel('Password:')
        self.layout.addWidget(self.password_label)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(self.password_input)

        self.confirm_password_label = QLabel("Confirm password:")
        self.layout.addWidget(self.confirm_password_label)

        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(self.confirm_password_input)

        self.lock_button = QPushButton('Block')
        self.lock_button.clicked.connect(self.check_password)
        self.layout.addWidget(self.lock_button)

        self.unlock_button = QPushButton('Unblock')
        self.unlock_button.clicked.connect(self.show_files)
        self.layout.addWidget(self.unlock_button)

        self.watermark = QLabel("Developed by Hubar")
        self.watermark.setAlignment(Qt.AlignCenter)
        self.watermark.setStyleSheet("color: #555; font-size: 14px;")
        self.layout.addWidget(self.watermark)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

    def update_selection_type(self):
        self.is_folder = self.folder_radio.isChecked()

    def choose_path(self):
        if self.is_folder:
            path = QFileDialog.getExistingDirectory(self, "Select folder")
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Select file")

        if path:
            self.path = path.replace('/', '\\')
            self.path_input.setText(self.path)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            dropped_path = urls[0].toLocalFile()
            self.path = dropped_path.replace('/', '\\')
            self.path_input.setText(self.path)
            self.is_folder = os.path.isdir(self.path)
            self.folder_radio.setChecked(self.is_folder)
            self.file_radio.setChecked(not self.is_folder)

    def check_password(self):
        try:
            entered_password = self.password_input.text()
            confirm_password = self.confirm_password_input.text()

            if not entered_password or not confirm_password:
                QMessageBox.critical(self, 'Error', "Password can not be empty!", QMessageBox.Ok, self)
            elif entered_password == confirm_password:
                self.password = entered_password
                self.lock_item()
                QMessageBox.information(self, 'Success', "Element is blocked!", QMessageBox.Ok, self)
                self.refresh_desktop()
            else:
                QMessageBox.critical(self, 'Error', "The passwords do not match!", QMessageBox.Ok, self)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"An error occurred: {str(e)}", QMessageBox.Ok, self)

    def lock_item(self):
        try:

            self.lock(self.path)
            vbs_file_path = self.create_vbs(self.path, self.password)
            shortcut_path = self.path + '.lnk'
            self.create_shortcut(vbs_file_path, shortcut_path)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"An error occurred: {str(e)}", QMessageBox.Ok, self)

    def create_vbs(self, path, password):
        home_dir = os.path.expanduser('~')
        name = path.split('\\')[-1]

        vbs_path = os.path.join(home_dir, f"{name}_lock.vbs")
        try:
            if self.is_folder:
                open_command = f"objShell.Explore \"{path}\""
            else:
                open_command = f"objShell.Open \"{path}\""

            vbs_content = f"""REM {path}\nDim sInput\nsInput = InputBox(\"Enter password\", \"Password required\")\nIf sInput = \"{password}\" Then\n    Set objShell = CreateObject(\"Shell.Application\")\n    {open_command}\nElse\n    MsgBox \"Incorrect password!!!\"\nEnd If"""

            with open(vbs_path, 'w') as f:
                f.write(vbs_content)

            self.lock(vbs_path)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"Failed to create VBS: {str(e)}", QMessageBox.Ok, self)
        return vbs_path

    def lock(self, path):
        try:
            if os.path.exists(path):
                call(['attrib', '+H', '+S', '+R', path])
            else:
                QMessageBox.warning(self, 'Attention', "Element to block not found.", QMessageBox.Ok, self)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"Failed to hide element: {str(e)}", QMessageBox.Ok, self)

    def create_shortcut(self, file_path, shortcut_path):
        try:
            shortcut = CoCreateInstance(
                shell.CLSID_ShellLink, None, CLSCTX_INPROC_SERVER, shell.IID_IShellLink
            )
            shortcut.SetPath(file_path)
            shortcut.SetDescription("Shortcut to block")
            shortcut.SetIconLocation("%SystemRoot%\\system32\\imageres.dll", 165)

            persist_file = shortcut.QueryInterface(IID_IPersistFile)
            persist_file.Save(shortcut_path, 0)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"Failed to create shortcut: {str(e)}", QMessageBox.Ok, self)

    def show_files(self):
        try:
            home_dir = os.path.expanduser('~')
            locked_files = [f for f in os.listdir(home_dir) if f.endswith('_lock.vbs')]

            if not locked_files:
                QMessageBox.information(self, 'Information', "No blocked elements found.", QMessageBox.Ok,
                                        self)
                return

            for file in locked_files:
                name = file.replace('_lock.vbs', '')
                unlock_confirmation = QMessageBox()
                unlock_confirmation.setWindowTitle("Unlock")
                unlock_confirmation.setText(f"Unblock: {name}?")
                unlock_confirmation.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                unlock_confirmation.setDefaultButton(QMessageBox.No)
                unlock_confirmation.setStyleSheet(
                    "QMessageBox { background-color: #3c3f41; color: #f0f0f0; } QPushButton { background-color: #555; color: #f0f0f0; } QPushButton:hover { background-color: #777; }"
                )
                unlock_confirmation.setWindowIcon(QIcon('icon.png'))
                if unlock_confirmation.exec() == QMessageBox.Yes:
                    self.unlock_item(file)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"An error occurred: {str(e)}", QMessageBox.Ok, self)

    def unlock_item(self, vbs_file):
        try:
            home_dir = os.path.expanduser('~')
            vbs_path = os.path.join(home_dir, vbs_file)

            if not os.path.exists(vbs_path):
                QMessageBox.warning(self, 'Attention', "Unlock file not found.", QMessageBox.Ok, self)
                return

            with open(vbs_path) as f:
                first_line = f.readline()
                path = first_line.replace('REM ', '').strip()

            if os.path.exists(path):
                call(['attrib', '-H', '-S', '-R', path])
            else:
                QMessageBox.warning(self, 'Attention', "Unlock item not found.", QMessageBox.Ok, self)

            call(['attrib', '-H', '-S', '-R', vbs_path])
            os.remove(vbs_path)

            shortcut_path = path + '.lnk'
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)

            QMessageBox.information(self, 'Success', f"Element {path} unlocked.", QMessageBox.Ok, self)
            self.refresh_desktop()
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"An error occurred: {str(e)}", QMessageBox.Ok, self)

    def refresh_desktop(self):
        try:
            desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            if os.path.exists(desktop_path):
                call(['attrib', '-H', desktop_path])
                call(['attrib', '+H', desktop_path])
        except Exception as e:
            QMessageBox.critical(self, 'Error', f"Failed to refresh desktop: {str(e)}", QMessageBox.Ok, self)


if __name__ == '__main__':
    app = QApplication([])
    app.setWindowIcon(QIcon('icon.png'))
    window = HubarLockApp()
    window.show()
    app.exec()
