from pathlib import Path
from xml.etree import ElementTree as ET
import shutil
import logging
from xml.dom import minidom
from PyQt5.QtCore import QObject, pyqtSignal
from PyQt5 import QtCore
import os
import stat
import glob
import pythoncom
from win32comext.shell import shell


class Signaller(QtCore.QObject):
    signal = QtCore.pyqtSignal(str, logging.LogRecord)


class QtHandler(logging.Handler):
    def __init__(self, slotfunc, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.signaller = Signaller()
        self.signaller.signal.connect(slotfunc)

    def emit(self, record):
        s = self.format(record)
        self.signaller.signal.emit(s, record)


class FileHandler(QObject):
    finished = pyqtSignal()
    Engine_status = pyqtSignal(str)

    clearing_progress = pyqtSignal(int)

    def __init__(self, logger):
        super(FileHandler, self).__init__()
        self.appsoftware_folder_path = ''
        self.ro_folder_path = ''
        self.before_path = ''
        self.red_path = ''
        self.ro_number = ''
        self.logger = logger
        self.total_removed_files_before = 0
        self.total_removed_files_red = 0
        self.hasher = QtCore.QCryptographicHash(QtCore.QCryptographicHash.Md5)

    def log(self, level, msg, sender=''):
        if sender == '':
            sender = self.__class__.__name__
        self.logger.log(level, msg, extra={'sender': sender})

    @QtCore.pyqtSlot(str, name='RO_Path')
    def ro_folder_path_changed(self, path):

        self.ro_folder_path = path
        self.ro_number = os.path.basename(path)
        self.before_path = self.ro_folder_path + '/Before'
        self.red_path = self.ro_folder_path + '/Red'

    @QtCore.pyqtSlot(str, name='Vault_Path')
    def appsoftware_folder_path_changed(self, path):
        self.appsoftware_folder_path = path

    def make_hash(self, file_path):
        self.hasher.reset()
        with open(file_path, 'rb') as fh:
            self.hasher.addData(fh.read())
        return bytes(self.hasher.result().toHex()).decode('UTF-8')

    def remove_empty_folders(self, base_path):
        counter = -1
        while counter != 0:
            counter = 0
            # Traversing through main path
            for root, dirs, files in os.walk(base_path):
                # Checking the size of tuple
                if not len(dirs) and not len(files):
                    # Removing the empty directory
                    os.rmdir(root)
                    counter += 1
                    self.log(logging.INFO, f'Empty folder {root} has been removed.')
                path = Path(root)
                parent_folder = path.parent.absolute()

                if not os.listdir(parent_folder) and (len(str(parent_folder)) >= len(str(base_path))):
                    os.rmdir(parent_folder)
                    counter += 1

                    self.remove_empty_folders(parent_folder.parent)

    def find_files_with_extension(self, root, ext):
        result = []
        for root, dirs, files in os.walk(root):
            xml_path = os.path.join(root, f'*.{ext}')
            result.extend(glob.glob(xml_path))
        return result

    def hwsx_deleter(self, xmls, element):
        counter = -1
        while counter != 0:
            counter = 0
            for xml in xmls:
                hafx_file_exist = False
                xml_folder = os.path.split(xml)[0]
                if os.path.exists(xml):
                    _xml = minidom.parse(xml)
                    nodes = _xml.getElementsByTagName(element)
                    for node in nodes:
                        path_in_hwsx = node.getAttribute('path')
                        complete_path = os.path.join(xml_folder, path_in_hwsx)
                        if os.path.exists(complete_path):
                            hafx_file_exist = True
                    if not hafx_file_exist:
                        os.chmod(xml, stat.S_IWRITE)
                        os.remove(xml)
                        counter += 1
                else:
                    continue

    def find_station_hwsx(self, topic):
        searching_folder_path = ''
        match topic.lower():
            case 'before':
                searching_folder_path = self.before_path
            case 'red':
                searching_folder_path = self.red_path
            case 'after':
                searching_folder_path = self.red_path.replace('red', 'after')
        return_list = []
        hwsx_files = self.find_files_with_extension(searching_folder_path, 'hwsx')
        for file in hwsx_files:
            if os.path.exists(file):
                _xml = minidom.parse(file)
                nodes = _xml.getElementsByTagName('workspaceItem')
                for node in nodes:
                    path_in_hwsx = node.getAttribute('path')
                    if 'Network' in path_in_hwsx:
                        return_list.append(file)
        return return_list

    def create_workspace(self, application_path_list, topic):
        fileguid_text = '239cfbc3-00f0-4212-9f6e-f22aa5798cd3'
        document = ET.Element('hidraw')
        hidrawworkspace = ET.SubElement(document, 'hidrawworkspace',
                                        {'version': '840', 'fileguid': fileguid_text})
        properties = ET.SubElement(hidrawworkspace, 'properties')
        ET.SubElement(properties, 'parameters')
        ET.SubElement(properties, 'extensionobjectsholders')
        workspace = ET.SubElement(hidrawworkspace, 'workspace')
        workspaceItems = ET.SubElement(workspace, 'workspaceItems')

        if topic.lower() == 'main':
            for item in application_path_list:
                ET.SubElement(workspaceItems, 'workspaceItem', {'type': 'workspace', 'path': f'{item}.hwsx'})
        elif topic.lower() == 'after':
            for item in application_path_list:
                item = item.replace('Red', 'After')
                ET.SubElement(workspaceItems, 'workspaceItem', {'type': 'workspace', 'path': f'{item}'})
        else:
            for item in application_path_list:
                ET.SubElement(workspaceItems, 'workspaceItem', {'type': 'workspace', 'path': f'{item}'})

        ET.SubElement(hidrawworkspace, 'views')
        et = ET.ElementTree(document)
        match topic.lower():
            case 'main':
                et.write(f'{self.ro_folder_path}/{self.ro_number}.hwsx', encoding='utf-8', xml_declaration=True)
            case 'after':
                et.write(f'{self.ro_folder_path}/After.hwsx', encoding='utf-8', xml_declaration=True)
            case 'before':
                et.write(f'{self.ro_folder_path}/Before.hwsx', encoding='utf-8', xml_declaration=True)
            case 'red':
                et.write(f'{self.ro_folder_path}/Red.hwsx', encoding='utf-8', xml_declaration=True)

    def clean_files(self):
        total_file_numbers = 0
        files_checked = 0
        for root, dirs, files in os.walk(self.before_path):
            for _ in files:
                total_file_numbers += 1

        for root, dirs, files in os.walk(self.before_path):
            for name in files:
                files_checked += 1
                percentage_checked = int((files_checked / total_file_numbers)*100)
                self.clearing_progress.emit(percentage_checked)

                before_file_path = os.path.join(root, name).replace("\\", "/")
                red_file_path = before_file_path.replace('Before', 'Red')
                before_hash_string = self.make_hash(before_file_path)
                try:
                    red_hash_string = self.make_hash(red_file_path)
                except:
                    red_hash_string = ''
                if before_hash_string == red_hash_string and not red_file_path.lower().endswith(
                        '.hwsx') and os.path.exists(
                    red_file_path):
                    try:
                        os.remove(before_file_path)
                        self.log(logging.INFO, f'No changes in file. File Removed : {before_file_path}')
                        self.total_removed_files_before += 1
                    except PermissionError:
                        self.log(logging.WARNING, f'Removing the read-only attribute from file : {before_file_path}')
                        os.chmod(before_file_path, stat.S_IWRITE)
                        os.remove(before_file_path)
                        self.log(logging.INFO, f'No changes in file. File Removed : {before_file_path}')
                    try:
                        os.remove(red_file_path)
                        self.log(logging.INFO, f'No changes in file. File Removed : {red_file_path}')
                        self.total_removed_files_red += 1
                    except PermissionError:
                        self.log(logging.WARNING, f'Removing the read-only attribute from file : {red_file_path}')
                        os.chmod(red_file_path, stat.S_IWRITE)
                        os.remove(red_file_path)
                        self.log(logging.INFO, f'No changes in file. File Removed : {red_file_path}')

        self.log(logging.INFO,
                 f'Cleaning files Finished. in total {self.total_removed_files_before} and {self.total_removed_files_red} files are removed from Before and Red Folders.')

    def shortcut_target(self, shortcutfile):
        pythoncom.CoInitialize()
        link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
                                          shell.IID_IShellLink)
        link.QueryInterface(pythoncom.IID_IPersistFile).Load(shortcutfile)
        target_path, _ = link.GetPath(shell.SLGP_UNCPRIORITY)
        return target_path

    def clean_shortcuts(self):
        counter = -1
        while counter != 0:
            counter = 0
            for root, dirs, files in os.walk(self.ro_folder_path):
                for name in files:
                    file_path = os.path.join(root, name).replace("\\", "/")
                    if file_path.lower().endswith('.lnk'):
                        target_of_link = self.shortcut_target(file_path)
                        if not os.path.exists(target_of_link):
                            os.remove(file_path)
                            counter += 1

    def create_modified_list(self):
        self.Engine_status.emit('create_modified_list_started')
        self.log(logging.INFO, f'Creating List of Modified Pages...')
        modified_pages_folder = f"{self.ro_folder_path}/Modifed Pages"

        os.makedirs(modified_pages_folder, exist_ok=True)

        for root, dirs, files in os.walk(self.before_path):
            # Open a file for writing affected pages in access mode 'a'
            for name in files:
                before_file_path = os.path.join(root, name).replace("\\", "/")
                red_file_path = before_file_path.replace('Before', 'Red')
                before_hash_string = self.make_hash(before_file_path)
                try:
                    red_hash_string = self.make_hash(red_file_path)
                except:
                    red_hash_string = ''
                if before_hash_string != red_hash_string and before_file_path.lower().endswith('.hgf'):
                    before_path_list = before_file_path.split('/')
                    before_index = before_path_list.index('Before') + 1
                    document_name_elements = before_path_list[before_index:-1]
                    document_name = '-'.join(document_name_elements)
                    modified_pages_path = f"{modified_pages_folder}/{document_name}.txt"

                    with open(modified_pages_path, "a") as f:
                        hgf_file_name = before_path_list[-1]
                        page_name = hgf_file_name.split('.')[0]
                        f.write(page_name + ',')
                        f.close()
        self.log(logging.INFO, f'Modified Pages List has been created.')
        self.Engine_status.emit('create_modified_list_done')

    def create_before_red_folders(self):
        self.Engine_status.emit('create_before_red_folders_started')
        self.log(logging.INFO, f'Copying Files to Before Folder...')

        current_folder_vault = os.path.basename(self.appsoftware_folder_path)

        shutil.copytree(self.appsoftware_folder_path, self.before_path + '/' + current_folder_vault, dirs_exist_ok=True)
        self.log(logging.INFO, f'Before folder created.')

        self.log(logging.INFO, f'Copying Files to Red Folder...')
        shutil.copytree(self.before_path, self.red_path, dirs_exist_ok=True)
        self.log(logging.INFO, f'Red folder created.')
        self.log(logging.INFO, f'Removing the Ready only attribute from red folder. {self.red_path}')
        for root, dirs, files in os.walk(self.red_path):
            for fname in files:
                full_path = os.path.join(root, fname)
                os.chmod(full_path, stat.S_IWRITE)
        self.log(logging.INFO, f'All Ready only attribute removed. {self.red_path}')
        self.Engine_status.emit('create_before_red_folders_done')

    def run(self):
        self.log(logging.INFO, f'Cleaning the unmodified files...')
        self.clean_files()

        self.log(logging.INFO, f'Removing All unused shortcut files...')
        self.clean_shortcuts()
        self.log(logging.INFO, f'All unused shortcut files has been removed.')

        self.log(logging.INFO, f'Removing All unused hwsx files...')
        self.hwsx_deleter(self.find_files_with_extension(self.ro_folder_path, 'hwsx'), 'workspaceItem')
        self.log(logging.INFO, f'All unused hwsx files has been removed.')

        self.log(logging.INFO, f'Removing Empty Folders...')
        self.remove_empty_folders(self.ro_folder_path)
        self.log(logging.INFO, f'All empty folders removed.')

        shutil.copytree(self.ro_folder_path + '/Red', self.ro_folder_path + '/After', dirs_exist_ok=True)

        self.log(logging.INFO, f'Creating the RO workspace...')

        self.create_workspace(['Before', 'After', 'Red'], 'main')

        self.create_workspace(self.find_station_hwsx('Before'), 'Before')
        self.create_workspace(self.find_station_hwsx('Red'), 'Red')
        self.create_workspace(self.find_station_hwsx('After'), 'After')

        self.log(logging.INFO, f'All workspaces are created.')
        self.finished.emit()
