# install pywin32 and Pyside2 module
# python -m pip install pywin32
# python -m pip install Pyside2

# TODO Implement Single Mutli-exrs functionality

import sys
import win32com.client
import os
from time import time

from PySide2 import QtCore, QtGui, QtWidgets
from psd_functions import *


class CreatePsbFromExrs(QtWidgets.QDialog):

    def __init__(self, parent=None):
        super(CreatePsbFromExrs, self).__init__(parent)
        # Window Variables

        self.setWindowTitle("Create PSB From EXRs")
        self.setMinimumWidth(600)
        self.height_single = 420
        self.height_multi = 420
        self.setFixedHeight(self.height_multi)
        self.multi_paths = []
        self.single_paths = []
        self.multi_browse_buttons = []
        self.single_browse_buttons = []
        self.number_paths = 5
        self.crypto_max = 99
        self.crypto_warning_threshold = 20
        
        # Icon
        icon_path = 'icon.ico'    
        icon = QtGui.QIcon(icon_path)
        self.setWindowIcon(icon)

        # Window Style
        font1 = QtGui.QFont()
        font1.setPointSize(10)
        font1.setBold(True)
        self.background_color = "background-color:rgb(137, 165, 209)"
        self.button_color = "background-color:rgb(176, 194, 233)"
        self.lineEdit_color = "background-color:rgb(255, 255, 255)"
        self.setStyleSheet(self.background_color)
        self.save_in_folder_text = "Save in Folder by default. Uncheck to select Output Folder"

        # CREATE WIDGETS
        def connect_factory(browse_function, editLine):
            def connect():
                browse_function(editLine)
            return connect

        for browse_type in ('multi', 'single'):
            for i in range(self.number_paths):
                exec('self.editLine_{0}_{1} = QtWidgets.QLineEdit()'.format(browse_type, str(i)))
                exec('self.editLine_{0}_{1}.setStyleSheet(self.lineEdit_color)'.format(browse_type, str(i)))
                exec('self.editLine_{0}_{1}.setFixedHeight(30)'.format(browse_type, str(i)))
                exec('self.editLine_{0}_{1}.setPlaceholderText("Enter Path")'.format(browse_type, str(i)))
                exec('self.button_browse_{0}_{1} = QtWidgets.QPushButton("Browse")'.format(browse_type, str(i)))
                exec('self.button_browse_{0}_{1}.setFixedWidth(100)'.format(browse_type, str(i)))
                exec('self.button_browse_{0}_{1}.setStyleSheet(self.button_color)'.format(browse_type, str(i)))
                if browse_type == 'multi':
                    exec('self.button_browse_{0}_{1}.clicked.connect(connect_factory(self.browse_folder, self.editLine_{0}_{1}))'.format(browse_type, str(i)))
                else:
                    exec('self.button_browse_{0}_{1}.clicked.connect(connect_factory(self.browse_file, self.editLine_{0}_{1}))'.format(browse_type, str(i)))
      
        self.button_check_multi = QtWidgets.QRadioButton("Multiple EXRs")
        self.button_check_single = QtWidgets.QRadioButton("Single Multi-Layered EXR")
        self.button_check_single.setVisible(True)
        self.button_check_multi.setChecked(True)
        self.button_options = QtWidgets.QButtonGroup()
        self.button_options.addButton(self.button_check_multi, id=1)
        self.button_options.addButton(self.button_check_single, id=2)

        # Multiple EXR Widgets
        self.label_multi_exr = QtWidgets.QLabel("PSB from multiple EXRs")
        self.label_multi_exr.setGeometry(QtCore.QRect(20, 10, 311, 20))
        self.label_multi_exr.setFont(font1)
        self.label_multi_exr.setSizePolicy(QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed))
        self.label_folder_path = QtWidgets.QLabel("Folder Path:")

        # self.editLine_folder = QtWidgets.QLineEdit()
        # self.editLine_folder.setStyleSheet(self.lineEdit_color)
        # self.editLine_folder.setPlaceholderText("Enter Folder containing EXRs")
        # self.editLine_folder.setFixedHeight(30)
        # self.button_browse_folder = QtWidgets.QPushButton("Browse")
        # self.button_browse_folder.setFixedWidth(100)
        # self.button_browse_folder.setStyleSheet(self.button_color)
        # self.multi_paths.append(self.editLine_folder)

        self.button_add_folder = QtWidgets.QPushButton("Add New Folder Path")
        self.button_add_folder.setStyleSheet(self.button_color)
        self.button_get_paths = QtWidgets.QPushButton("Print Paths")
        self.checkbox_subfolder = QtWidgets.QCheckBox("Include Subfolders")
        self.checkbox_cryptomatte1 = QtWidgets.QCheckBox("Include Cryptomatte")

        # Single EXR Widgets
        self.label_single_exr = QtWidgets.QLabel("PSB from single multi-layered EXR")
        self.label_single_exr.setFont(font1)
        self.label_single_exr.setSizePolicy(QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed))
        self.label_file_path = QtWidgets.QLabel("File Path:")
        # self.editLine_file = QtWidgets.QLineEdit()
        # self.editLine_file.setStyleSheet(self.lineEdit_color)
        # self.editLine_file.setPlaceholderText("Enter multi-layered EXR File")
        # self.editLine_file.setFixedHeight(30)
        # self.single_paths.append(self.editLine_file)

        self.button_browse_file = QtWidgets.QPushButton("Browse File")
        self.button_browse_file.setStyleSheet(self.button_color)
        self.button_add_file = QtWidgets.QPushButton("Add New File Path")
        self.button_add_file.setStyleSheet(self.button_color)
        self.button_get_file = QtWidgets.QPushButton("Print Files")
        self.checkbox_cryptomatte2 = QtWidgets.QCheckBox("Include Cryptomatte")

        # Outputs
        self.label_output = QtWidgets.QLabel("Outputs")
        self.label_output.setFont(font1)
        self.label_output.setSizePolicy(QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed))
        self.label_output_folder = QtWidgets.QLabel("Output Folder:")
        self.editLine_output = QtWidgets.QLineEdit()
        self.editLine_output.setStyleSheet(self.lineEdit_color)
        self.editLine_output.setPlaceholderText(self.save_in_folder_text)
        self.editLine_output.setEnabled(False)
        self.button_browse_output = QtWidgets.QPushButton("Browse")
        self.button_browse_output.setFixedWidth(100)
        self.button_browse_output.setStyleSheet(self.button_color)
        self.bit_list = QtWidgets.QComboBox()
        self.bit_list.setStyleSheet(self.lineEdit_color)
        self.checkbox_save_in_folder = QtWidgets.QCheckBox("Save in Folder Path")
        self.checkbox_save_in_folder.setChecked(True)
        self.button_cryptomatte = QtWidgets.QPushButton("Check Cryptomatte Number of Objects")
        self.button_cryptomatte.setStyleSheet(self.button_color)
        self.button_create_psb = QtWidgets.QPushButton("CREATE PSB FROM EXRS")
        self.button_create_psb.setStyleSheet(self.button_color)
        self.button_create_psb.setFont(font1)
        self.checkbox_close_psb = QtWidgets.QCheckBox("Close PSBs after saving")
        self.checkbox_close_psb.setChecked(True)
        self.checkbox_open_dirs = QtWidgets.QCheckBox("Open Folder Directories when finished")
        self.checkbox_open_dirs.setChecked(True)
        
        self.bit_list.insertItem(0, "8 bit Mode")
        self.bit_list.insertItem(1, "16 bit Mode")
        self.bit_list.insertItem(2, "32 bit Mode")

        # Main Layout
        self.main_layout = QtWidgets.QVBoxLayout()
        self.setLayout(self.main_layout)
        
        # Layout EXR Option
        self.layout_exr_option = QtWidgets.QGridLayout()
        self.layout_exr_option.addWidget(self.button_check_multi, 1, 0)
        self.layout_exr_option.addWidget(self.button_check_single, 1, 1)

        # Layout Multiple EXRs
        self.layout_multi_exr = QtWidgets.QGridLayout()
        self.layout_multi_exr.addWidget(self.label_multi_exr, 1, 0)
        self.layout_multi_exr.addWidget(self.label_folder_path, 2, 0)
        self.add_path(1)
        # self.layout_multi_exr.addWidget(self.editLine_folder, 3, 0)
        # self.layout_multi_exr.addWidget(self.button_browse_folder, 3, 1)
        # self.layout_multi_exr.setColumnMinimumWidth(1, 2)

        self.layout_multi_exr_options = QtWidgets.QGridLayout()
        
        self.layout_multi_exr_options.addWidget(self.button_add_folder, 1, 0, 1, 2)
        # self.layout_multi_exr_options.addWidget(self.checkbox_subfolder, 2, 0)
        self.layout_multi_exr_options.addWidget(self.checkbox_cryptomatte1, 2, 0)
        
        # Layout Single EXR
        self.layout_single_exr = QtWidgets.QGridLayout()
        self.layout_single_exr.addWidget(self.label_single_exr, 1, 0)
        self.layout_single_exr.addWidget(self.label_file_path, 2, 0)
        self.add_path(2)
        # self.layout_single_exr.addWidget(self.editLine_file)
        self.layout_single_exr_options = QtWidgets.QGridLayout()
        # self.layout_single_exr_options.addWidget(self.button_browse_file, 1, 0)
        self.layout_single_exr_options.addWidget(self.button_add_file, 1, 0, 1, 2)
        self.layout_single_exr_options.addWidget(self.checkbox_cryptomatte2, 2, 0)

        # Layout Outputs
        self.layout_output = QtWidgets.QVBoxLayout()
        self.layout_output.addWidget(self.label_output, QtCore.Qt.AlignTop)
        self.layout_output.addWidget(self.label_output_folder)
        self.layout_output_browse = QtWidgets.QGridLayout()
        self.layout_output_browse.addWidget(self.editLine_output, 1, 0)
        self.layout_output_browse.addWidget(self.button_browse_output, 1, 1)
        
        self.layout_output_options = QtWidgets.QGridLayout()
        self.layout_output_options.addWidget(self.bit_list, 1, 0)
        self.layout_output_options.addWidget(self.checkbox_save_in_folder, 1, 1)
        self.layout_output_options.addWidget(self.checkbox_close_psb, 2, 0)
        self.layout_output_options.addWidget(self.checkbox_open_dirs, 2, 1)

        # Add to MainLayout
        self.main_layout.addLayout(self.layout_exr_option)
        self.main_layout.addLayout(self.layout_multi_exr)
        self.main_layout.addLayout(self.layout_multi_exr_options)
        self.main_layout.addLayout(self.layout_single_exr)
        self.main_layout.addLayout(self.layout_single_exr_options)
        self.main_layout.addLayout(self.layout_output)
        self.main_layout.addLayout(self.layout_output_browse)
        self.main_layout.addLayout(self.layout_output_options) 
        self.main_layout.addWidget(self.button_cryptomatte)
        self.main_layout.addWidget(self.button_create_psb)
        self.toogle_layout()
        
        
        # Connections Multi EXR
        # self.button_browse_multi_0.clicked.connect(lambda: self.browse_folder(self.editLine_multi_0))
        # self.button_browse_folder.clicked.connect(self.browse_folder)
        self.button_cryptomatte.clicked.connect(self.get_cryptomatte_info)
        self.button_create_psb.clicked.connect(self.create_psb)
        self.button_add_folder.clicked.connect(lambda: self.add_path(1))
        self.button_get_paths.clicked.connect(self.print_paths)
        self.button_add_file.clicked.connect(lambda: self.add_path(2))
        self.button_options.idToggled.connect(self.toogle_layout)
        self.button_browse_file.clicked.connect(self.browse_file)
        self.button_browse_output.clicked.connect(connect_factory(self.browse_folder, self.editLine_output))
        self.checkbox_save_in_folder.stateChanged.connect(self.save_in_folder)
        self.check_exr_io()
    


    # Window Functions

    def check_exr_io(self):
        if not exr_io_check():
            self.show_message(["Please install the exr-io plugin v 2.00 or higher."])
            sys.exit(app.exec_())


    def save_in_folder(self):
        if self.checkbox_save_in_folder.isChecked():
            self.editLine_output.setEnabled(False)
            self.editLine_output.setText(self.save_in_folder_text)
        else:
            self.editLine_output.setEnabled(True)

    def toogle_layout(self):
        if self.button_options.checkedId() == 1:
            on = 0
            off = 1
            self.setFixedHeight(self.height_multi)
        else:
            on = 1
            off = 0
            self.setFixedHeight(self.height_single)
        def layout_state(layout, state):
            for i in range(layout.count()):
                item = layout.itemAt(i)
                item.widget().setVisible(state)
        layouts = (self.layout_multi_exr, self.layout_single_exr)
        options = (self.layout_multi_exr_options, self.layout_single_exr_options)
        layout_state(layouts[on], True)
        layout_state(options[on], True)
        layout_state(layouts[off], False)
        layout_state(options[off], False)


    def print_paths(self):
        # for path in self.paths:
            print("yes")
            

    def add_path(self, index):
        if index == 1:
            number = len(self.multi_paths)
            if number >= self.number_paths:
                return
            layout = self.layout_multi_exr            
            row = number + 3
            exec('self.layout_multi_exr.addWidget(self.editLine_multi_{0}, {1}, 0)'.format( number, row))
            exec('self.layout_multi_exr.addWidget(self.button_browse_multi_{0}, {1}, 1)'.format(number, row))
            exec('self.multi_paths.append(self.editLine_multi_{0})'.format(number))
            exec('self.multi_browse_buttons.append(self.button_browse_multi_{0})'.format(number))
            self.height_multi += 32
            self.setFixedHeight(self.height_multi)

        elif index == 2:
            number = len(self.single_paths)
            if number >= self.number_paths:
                return
            layout = self.layout_single_exr
            row = number + 4
            exec('self.layout_single_exr.addWidget(self.editLine_single_{0}, {1}, 0)'.format(number, row))
            exec('self.layout_single_exr.addWidget(self.button_browse_single_{0}, {1}, 1)'.format(number, row))
            exec('self.single_paths.append(self.editLine_single_{0})'.format(number))
            exec('self.multi_browse_buttons.append(self.button_browse_single_{0})'.format(number))
            self.height_single += 32
            self.setFixedHeight(self.height_single)


    def browse_folder(self, lineEdit):
        last_file = os.environ.get('LAST_PATH')
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, ("Open Directory"), last_file, QtWidgets.QFileDialog.ShowDirsOnly)
        if directory != "":
            lineEdit.setText(directory)
            os.environ['LAST_PATH'] = os.path.dirname(directory)


    def browse_file(self, lineEdit):
        last_file = os.environ.get('LAST_PATH')
        filename, selected = QtWidgets.QFileDialog.getOpenFileName(self, ("Single multi-layered EXRs"), last_file, ("EXR (*.exr)"))
        if selected:
            lineEdit.setText(filename)
            os.environ['LAST_PATH'] = os.path.dirname(filename)

            
    def get_directories(self):
        errors = []
        directories = []
        for path in self.multi_paths:
            path_str = r'{}'.format(str(path.text()))
            if os.path.exists(path_str):
                if path_str not in directories:
                    if not self.checkbox_subfolder.isChecked():
                        directories.append(path_str)
                    else:
                        for dirpath, dirnames, filenames in os.walk(path_str):
                            if not dirnames:
                                directories.append(dirpath) 
            elif path_str == "":
                pass
            else:
                errors.append("{} is not a valid directory".format(path_str))
        return directories, errors


    def get_files(self):
        errors = []
        files = []
        for path in self.single_paths:
            path_str = r'{}'.format(str(path.text()))
            if os.path.exists(path_str):
                if path_str not in files:
                        files.append(path_str)
            elif path_str == "":
                pass
            else:
                errors.append("- {} is not a valid directory".format(path_str))
        return files, errors

        
    def show_message(self, message):
        total_message = "\n".join(message)
        msgBox = QtWidgets.QMessageBox()
        msgBox.setText(total_message)
        msgBox.setStandardButtons(QtWidgets.QMessageBox.Close)
        msg_action = msgBox.exec_()

    
    def get_paths(self):
        valid_directories = []
        if self.button_options.checkedId() == 1:
            func = self.get_directories
        else:
            func = self.get_files

        directories, errors = func()

        if len(errors) > 0:
            self.show_message(errors)
            return False, directories
        return True, directories


    def get_cryptomatte_info(self):
        run, directories = self.get_paths()
        crypto_message = []
        if run:
            if self.button_options.checkedId() == 1:
                for directory in directories:
                    files = os.listdir(directory)
                    for _file in files:
                        if "crypto" in _file:
                            full_path = os.path.join(directory, _file)
                            name, extension = os.path.splitext(full_path)
                            if extension == ".exr":
                                crypto_objects = read_crypto_elements(full_path)
                                _message = "There are {} crypto layers in {}".format(len(crypto_objects), _file)
                                crypto_message.append(_message)

                if len(crypto_message) == 0:
                    self.show_message(["There are no crypto files"])
                else:
                    self.show_message(crypto_message)
            else:
                # TODO single_exr
                pass


    def check_cryptomatte(self):
        run, directories = self.get_paths()
        fail = False
        run_crypto = True
        crypto_message = []
        if self.button_options.checkedId() == 1:
            for directory in directories:
                files = os.listdir(directory)
                for _file in files:
                    if "crypto" in _file:
                        full_path = os.path.join(directory, _file)
                        name, extension = os.path.splitext(full_path)
                        if extension == ".exr":
                            crypto_objects = read_crypto_elements(full_path)
                            crypto_length = len(crypto_objects)
                            if crypto_length > self.crypto_max: 
                                _message = "Can't run {} crypto layers in {}".format(crypto_length, _file)
                                fail = True
                                crypto_message.append(_message)
                            elif self.crypto_warning_threshold < crypto_length < self.crypto_max: 
                                _message = "There are {} crypto layers in {}".format(crypto_length, _file)
                                crypto_message.append(_message)
                                run_crypto = False

        if len(crypto_message) > 0:
            crypto_message.append("What whould you like to do?")
            total_message = "\n".join(crypto_message)
            msgBox = QtWidgets.QMessageBox()
            msgBox.setText(total_message)
            crypto_off = msgBox.addButton("Skip Cryptomatte", QtWidgets.QMessageBox.ActionRole)
            if not fail:
               proceed = msgBox.addButton("Proceed", QtWidgets.QMessageBox.ActionRole)
            msgBox.setStandardButtons(QtWidgets.QMessageBox.Cancel)
            msgBox.exec()

            if msgBox.clickedButton() == crypto_off:
                self.checkbox_cryptomatte1.setChecked(False)
                run_crypto = True
            elif msgBox.clickedButton() == proceed:
                run_crypto = True
        

        return fail or run_crypto


    def create_psb(self):
        run, directories = self.get_paths()
        if run:
            if not self.check_cryptomatte():
                return
                
            start_time = time()
            bit_values = [8, 16, 32]
            bit_value = bit_values[self.bit_list.currentIndex()]
            crypto_state = self.checkbox_cryptomatte1.isChecked()

            psbs_created = ["Psb Files Created:"]

            for folder in directories:
                exr_files = []
                exr_folder = r'{}'.format(folder)
                
                # Get output folder
                if self.checkbox_save_in_folder.isChecked():
                    output_folder = exr_folder
                else:
                    if os.path.exists(self.editLine_output.text()):
                        output_folder = self.editLine_output.text()
                    else:
                        self.show_message(["Please select an existing output folder."])
                        return

                # Get exr files
                for root, dirs, files in os.walk(exr_folder):
                    for filename in files:
                        name, extension = os.path.splitext(filename)
                        if extension == ".exr":
                            full_path = os.path.join(exr_folder, filename)
                            exr_files.append({"path":full_path, "name":name})

                if len(exr_files) == 0:
                    self.show_message("Please select a directory containing EXRs.")
                    return

                try:
                    psb = create_multexr_psb(exr_files, output_folder, crypto_state, bit_value)
                    psbs_created.append(psb)
                except:
                    psbs_created.append("There was an error runntin {}.".format(folder))


            total_time = round((time() - start_time), 2)
            psbs_created.append("Total Processing time: {} seconds".format(total_time))
            self.show_message(psbs_created)
            print(directories)
        # sys.exit(app.exec_())

    

if __name__ =="__main__":
    # Create the Qt Application
    app = QtWidgets.QApplication(sys.argv)

    window = CreatePsbFromExrs()
    window.show()
    # Run the main Qt loop
    sys.exit(app.exec_())

