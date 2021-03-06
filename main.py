import os
import shutil
import sys
import tempfile
import webbrowser
import zipfile
import requests
import ctypes

from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal
from win32com.client import Dispatch

from ui import Ui_MainWindow


def open_privacy_policy():
    webbrowser.open('https://sites.google.com/view/edu-area/политика-конфиденциальности')


class Setup(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon(resource_path('logo.ico')))
        self.logo_label.setPixmap(QPixmap(resource_path('logo.ico')))
        self.privacy_policy_btn.clicked.connect(open_privacy_policy)
        self.path = 'C:/Program Files'
        self.shell = None
        self.installation = Installation(self)
        self.path_le.setReadOnly(True)
        self.path_le.setText(self.path)
        self.change_path_btn.clicked.connect(self.change_path)
        self.install_btn.clicked.connect(self.install)
        self.installation.installed.connect(self.create_shortcut)

    def install(self):
        self.state_name.setText('Идёт установка, пожалуйста, подождите...')
        self.repaint()
        self.installation.start()

    def create_shortcut(self):
        desktop = os.environ['USERPROFILE'] + '/Desktop'
        path = os.path.join(desktop, "AREA-Student.lnk")
        target = self.path + '/AREA-Student/AREA-Student.exe'
        working_dir = self.path + '/AREA-Student'
        icon = self.path + '/AREA-Student/System Files/logo.ico'
        self.shell = Dispatch('WScript.Shell')
        shortcut = self.shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = working_dir
        shortcut.IconLocation = icon
        shortcut.save()
        # run as admin flag
        # with open(path, "rb") as f2:
        #     ba = bytearray(f2.read())
        # ba[0x15] |= 0x20
        # with open(path, "wb") as f3:
        #     f3.write(ba)
        self.close()

    def change_path(self):
        new_path = QFileDialog().getExistingDirectory()
        self.path = new_path if new_path else self.path
        self.path_le.setText(self.path)


class Installation(QThread):
    installed = pyqtSignal()

    def __init__(self, parent):
        super(Installation, self).__init__()
        self.parent = parent

    def run(self):
        if os.path.exists(self.parent.path + '/AREA-Student'):
            shutil.rmtree(self.parent.path + '/AREA-Student')
        os.mkdir(self.parent.path + '/AREA-Student')
        self.parent.path = self.parent.path_le.text()
        version = self.parent.system.currentText()
        link = 'https://github.com/AREA-team/AREA-Student/releases/download/v1.0/' \
               'AREA-Student.32bit.zip'
        if version == 'Windows 32 bit':
            link = 'https://github.com/AREA-team/AREA-Student/releases/download/v1.0/' \
                   'AREA-Student.32bit.zip'
        elif version == 'Windows 64 bit':
            link = 'https://github.com/AREA-team/AREA-Student/releases/download/v1.0/' \
                   'AREA-Student.zip'
        response = requests.get(link)
        file = tempfile.TemporaryFile()
        file.write(response.content)
        fzip = zipfile.ZipFile(file)
        fzip.extractall(self.parent.path + '/AREA-Student')
        file.close()
        fzip.close()
        if version == 'Windows 32 bit':
            os.chdir(self.parent.path + '/AREA-Student')
            os.rename('AREA-Student 32bit.exe', 'AREA-Student.exe')
        self.installed.emit()
        self.quit()


def resource_path(relative):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative)
    else:
        return os.path.join(os.path.abspath(""), relative)


def is_admin():
    return ctypes.windll.shell32.IsUserAnAdmin()


if __name__ == '__main__':
    if is_admin():
        app = QApplication(sys.argv)
        wnd = Setup()
        wnd.show()
        sys.exit(app.exec())
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
