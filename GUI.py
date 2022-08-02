from datetime import datetime
import logging
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QFileDialog
from PyQt5.uic import loadUi
import sys
from PyQt5.QtWidgets import QApplication
import os
from config_manager import DataSetting
from Engine import FileHandler, QtHandler


class MainWindow(QtWidgets.QMainWindow):
    ro_folder_signal = pyqtSignal(str)
    appsoftware_folder_signal = pyqtSignal(str)
    clean_files_signal = pyqtSignal()
    create_modified_list_signal = pyqtSignal()
    create_before_red_signal = pyqtSignal()
    # handler_busy = pyqtSignal()

    COLORS = {
        logging.DEBUG: 'black',
        logging.INFO: 'blue',
        logging.WARNING: 'orange',
        logging.ERROR: 'red',
        logging.CRITICAL: 'purple',
    }

    def __init__(self, parent=None):
        super().__init__(parent)

        self.ro_list = []
        self.red_folder_path = ''
        self.before_folder_path = ''
        self.vault_appsoftware_folder_path = ''
        self.ro_number = ''
        self.ro_folder_path = ''
        self.thread = QThread()
        self.data_setting = DataSetting()
        self.current_working_folder_path = self.data_setting.working_folder_path

        self.setup_ui()

        # Logger Setup
        log_file_name = f".\logs\ROMaker_{datetime.now().strftime('%Y_%m_%d_%H_%M')}.log"
        self.log_file_handler = logging.FileHandler(log_file_name)
        self.log_file_handler.setFormatter(
            logging.Formatter('{asctime:<10s}   {sender:<14s}   {levelname:<9s}    {message}', style='{'))

        self.logger = logging.getLogger('ROLogs')
        self.logger.setLevel(logging.INFO)
        self.log_handler = QtHandler(self.update_status)
        self.log_handler.setFormatter(
            logging.Formatter('{asctime:<10s}   {sender:<14s}   {levelname:<9s}    {message}', style='{'))
        self.log_handler.setLevel(logging.DEBUG)
        self.logger.addHandler(self.log_handler)
        self.logger.addHandler(self.log_file_handler)

        self.log(logging.DEBUG, f'UI Started')
        f = QtGui.QFont('nosuchfont')
        f.setStyleHint(f.Monospace)
        self.logger_box.setFont(f)
        self.logger_box.setReadOnly(True)

        self.handler = FileHandler(self.logger)
        self.handler.moveToThread(self.thread)
        self.ro_folder_signal.connect(self.handler.ro_folder_path_changed)
        self.appsoftware_folder_signal.connect(self.handler.appsoftware_folder_path_changed)

        self.create_before_red_signal.connect(self.handler.create_before_red_folders)

        self.thread.finished.connect(self.thread.quit)
        self.thread.finished.connect(self.thread.deleteLater)
        self.clean_files_signal.connect(self.handler.run)
        self.create_modified_list_signal.connect(self.handler.create_modified_list)

        self.handler.Engine_status.connect(self.file_handler_status_changed)
        self.handler.clearing_progress.connect(self.cleaning_percentage_changed)

        self.thread.start()

    def setup_ui(self):
        loadUi(".\GUI\GUI.ui", self)
        self.lineEdit_RoNumber.setEnabled(False)
        self.btn_BrowseAppsoftware.setEnabled(False)
        self.btn_CreateBeforeRed.setEnabled(False)
        self.btn_OpenExplorer_Create.setEnabled(False)
        self.btn_OpenExplorer_Edit.setEnabled(False)

        self.setWindowIcon(QtGui.QIcon('Resources\HE.png'))
        self.btn_cleanFiles.setEnabled(False)
        self.btn_CreateModified.setEnabled(False)

        self.btn_BrowseWF.clicked.connect(self.browse_working_directory)
        self.btn_BrowseAppsoftware.clicked.connect(self.browse_appsoftware_folder)
        self.btn_CreateBeforeRed.clicked.connect(self.create_before_red_folders)
        self.btn_exit.clicked.connect(lambda: self.close())
        self.btn_cleanFiles.clicked.connect(self.clean_files)
        self.btn_CreateModified.clicked.connect(self.create_list)

        self.rbtn_create_ro.toggled.connect(self.radiobutton_toggled)
        self.rbtn_create_ro.toggled.connect(self.lineEdit_RoNumber.clear)
        self.rbtn_create_ro.toggled.connect(self.lineEdit_Appsoftware.clear)

        self.btn_OpenExplorer_Create.clicked.connect(self.open_file_explorer)
        self.btn_OpenExplorer_Edit.clicked.connect(self.open_file_explorer)

        self.update_ro_number_list()
        if len(self.ro_list):
            self.rbtn_edit_ro.setChecked(True)
        else:
            self.rbtn_create_ro.setChecked(True)
        self.radiobutton_toggled()

        self.ComboBox_SelectRO.currentTextChanged.connect(self.on_ro_number_selected)
        self.lineEdit_RoNumber.textChanged.connect(self.on_ro_number_text_changed)
        self.lineEdit_WorkingFolder.setText(self.data_setting.working_folder_path)

    @QtCore.pyqtSlot(int)
    def cleaning_percentage_changed(self, percentage):
        self.progressBar.setValue(percentage)

    @QtCore.pyqtSlot(str)
    def file_handler_status_changed(self, status_text):
        match status_text.lower():
            case 'create_modified_list_started':
                self.btn_CreateModified.setEnabled(False)
                self.btn_cleanFiles.setEnabled(False)
            case 'create_modified_list_done':
                self.btn_CreateModified.setEnabled(True)
                self.btn_cleanFiles.setEnabled(True)
            case 'create_before_red_folders_started':
                self.btn_BrowseAppsoftware.setEnabled(False)
                self.btn_CreateBeforeRed.setEnabled(False)

            case 'create_before_red_folders_done':
                self.btn_BrowseAppsoftware.setEnabled(True)
                self.btn_CreateBeforeRed.setEnabled(True)
                self.btn_OpenExplorer_Create.setEnabled(True)

    def log(self, level, msg, sender=''):
        if sender == '':
            sender = self.__class__.__name__
        self.logger.log(level, msg, extra={'sender': sender})

    def update_ro_number_list(self):
        self.ComboBox_SelectRO.clear()
        try:
            self.ro_list = os.listdir(self.current_working_folder_path)
            if len(self.ro_list):
                self.ComboBox_SelectRO.setEnabled(True)
                self.ComboBox_SelectRO.addItem('')
                for item in self.ro_list:
                    self.ComboBox_SelectRO.addItem(item)
            else:
                self.ComboBox_SelectRO.setEnabled(False)
        except FileNotFoundError:
            pass

    @QtCore.pyqtSlot(str, logging.LogRecord)
    def update_status(self, status, record):
        color = self.COLORS.get(record.levelno, 'black')
        s = '<pre><font color="%s">%s</font></pre>' % (color, status)
        self.logger_box.appendHtml(s)

    @property
    def working_folder_path(self):
        return self.data_setting.working_folder_path

    @working_folder_path.setter
    def working_folder_path(self, path):
        self.data_setting.working_folder_path = path

    @QtCore.pyqtSlot()
    def on_ro_number_text_changed(self):
        self.ro_number = self.lineEdit_RoNumber.text()
        self.ro_folder_path = self.working_folder_path + '/' + self.ro_number
        self.ro_folder_signal.emit(self.ro_folder_path)
        if self.vault_appsoftware_folder_path and self.ro_number:
            self.btn_CreateBeforeRed.setEnabled(True)
        else:
            self.btn_CreateBeforeRed.setEnabled(False)

    def browse_appsoftware_folder(self):
        self.vault_appsoftware_folder_path = QFileDialog.getExistingDirectory(self, 'Select Project Appsoftware Folder',
                                                                              'C:\Vault')
        if self.vault_appsoftware_folder_path and self.ro_number:
            self.btn_CreateBeforeRed.setEnabled(True)
        else:
            self.btn_CreateBeforeRed.setEnabled(False)

        self.lineEdit_Appsoftware.setText(self.vault_appsoftware_folder_path)
        self.appsoftware_folder_signal.emit(self.vault_appsoftware_folder_path)

    def browse_working_directory(self):
        self.current_working_folder_path = QFileDialog.getExistingDirectory(self, 'Set Working Folder',
                                                                            self.current_working_folder_path)
        if self.current_working_folder_path:
            self.lineEdit_WorkingFolder.setText(self.current_working_folder_path)
            self.data_setting.working_folder_path = self.current_working_folder_path
            self.update_ro_number_list()

    def create_before_red_folders(self):
        self.create_before_red_signal.emit()

    def radiobutton_toggled(self):
        if self.rbtn_create_ro.isChecked():
            self.Frame_Create_RO.show()
            self.Frame_Edit_RO.hide()
            self.lineEdit_RoNumber.setEnabled(True)
            self.btn_BrowseAppsoftware.setEnabled(True)
            self.btn_CreateBeforeRed.setEnabled(False)
            self.btn_OpenExplorer_Create.setEnabled(False)

        else:
            self.Frame_Create_RO.hide()
            self.Frame_Edit_RO.show()
            self.ComboBox_SelectRO.setEnabled(True)
            self.update_ro_number_list()

    def on_ro_number_selected(self, text):
        if text != '':
            self.ro_folder_path = self.current_working_folder_path + '/' + text
            self.ro_folder_signal.emit(self.ro_folder_path)
            self.btn_cleanFiles.setEnabled(True)
            self.btn_CreateModified.setEnabled(True)
            self.btn_OpenExplorer_Edit.setEnabled(True)
        else:
            self.btn_cleanFiles.setEnabled(False)
            self.btn_CreateModified.setEnabled(False)
            self.btn_OpenExplorer_Edit.setEnabled(False)

    def create_list(self):
        self.create_modified_list_signal.emit()

    def clean_files(self):
        self.clean_files_signal.emit()

    def open_file_explorer(self):
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(self.ro_folder_path))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
