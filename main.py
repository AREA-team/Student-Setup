import os
import shutil
import sys
import tempfile
import webbrowser
import zipfile
import requests
import ctypes

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
        self.privacy_policy_btn.clicked.connect(open_privacy_policy)
        self.path = 'C:/Program Files'
        self.shell = None
        self.path_le.setReadOnly(True)
        self.path_le.setText(self.path)
        self.change_path_btn.clicked.connect(self.change_path)
        self.installation = Installation(self)
        self.install_btn.clicked.connect(self.installation.start)
        self.installation.installed.connect(self.create_shortcut)

    def create_shortcut(self):
        os.rename(self.path + '/AREA-Student-1.0-beta',
                  self.path + '/AREA-Student')
        desktop = os.environ['USERPROFILE'] + '/Desktop'
        path = os.path.join(desktop, "AREA-Student.lnk")
        target = self.path + '/AREA-Student/AREA-Student.exe'
        working_dir = self.path + '/AREA-Student'
        icon = self.path + '/AREA-Student/System Files/Logo.ico'
        self.shell = Dispatch('WScript.Shell')
        shortcut = self.shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = working_dir
        shortcut.IconLocation = icon
        shortcut.save()
        # run as admin flag
        with open(path, "rb") as f2:
            ba = bytearray(f2.read())
        ba[0x15] |= 0x20
        with open(path, "wb") as f3:
            f3.write(ba)
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
        self.parent.state_name.setText('Идёт установка, пожалуйста, подождите...')
        self.parent.repaint()
        self.parent.path = self.parent.path_le.text()
        if os.path.exists(self.parent.path + '/AREA-Student-1.0-beta'):
            shutil.rmtree(self.parent.path + '/AREA-Student-1.0-beta')
        response = requests.get('https://github.com/AREA-team/AREA-Student/archive/v1.0-beta.zip')
        file = tempfile.TemporaryFile()
        file.write(response.content)
        fzip = zipfile.ZipFile(file)
        fzip.extractall(self.parent.path)
        file.close()
        fzip.close()
        self.installed.emit()
        self.quit()


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
