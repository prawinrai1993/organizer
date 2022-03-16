import sys
import os
import ntpath
import shutil
import zipfile
import re
from os import path
import pandas as pd
from PySide6 import QtCore, QtWidgets, QtGui
from datetime import date
from PySide6.QtWidgets import  QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QLineEdit, QMessageBox, QFileDialog, QProgressBar


DEBUG = True

def log(string):
    if DEBUG:
        print(string)

class OrganizerApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.progressBar = QProgressBar()
        self.sourceLabel = QLabel("source path: ")
        self.destinationLabel = QLabel("destination path: ")
        self.sourceLineEdit = QLineEdit()
        self.destinationLineEdit = QLineEdit()
        self.sourceButton = QPushButton("select source path")
        self.destinationButton = QPushButton("select destination path")
        self.runButton = QPushButton("run organizer")
        self.vLayout = QVBoxLayout(self)
        self.hLayout1 = QHBoxLayout(self)
        self.hLayout2 = QHBoxLayout(self)
        self.hLayout3 = QHBoxLayout(self)
        self.hLayout4 = QHBoxLayout(self)

        self.hLayout1.addWidget(self.sourceLabel)
        self.hLayout1.addWidget(self.sourceLineEdit)
        self.hLayout1.addWidget(self.sourceButton)
        self.hLayout2.addWidget(self.destinationLabel)
        self.hLayout2.addWidget(self.destinationLineEdit)
        self.hLayout2.addWidget(self.destinationButton)
        self.hLayout3.addWidget(self.runButton)
        self.hLayout4.addWidget(self.progressBar)

        self.vLayout.addLayout(self.hLayout1)
        self.vLayout.addLayout(self.hLayout2)
        self.vLayout.addLayout(self.hLayout3)
        self.vLayout.addLayout(self.hLayout4)

        self.runButton.clicked.connect(self.runbuttonClicked)
        self.sourceButton.clicked.connect(self.sourcebuttonClicked)
        self.destinationButton.clicked.connect(self.destinationbuttonClicked)

    @QtCore.Slot()
    def runOrganizer(self):
        print("running organizer")
        sourcefilespath = self.sourceLineEdit.text()
        destinationfilespath = self.destinationLineEdit.text()
        filepathList = []
        filenameList = []
        result = []
        excelData = []

        excelData.append(["Source","Destination"])

        filecount = 0
        progresscount = 0
        self.progressBar.setValue(0)

        # loop for unzipping files
        for r, d, f in os.walk(sourcefilespath):
            for file in f:
                if ".zip" in file:
                    folderPath = os.path.dirname(os.path.abspath(os.path.join(r, file)))
                    with zipfile.ZipFile(os.path.join(r, file), 'r') as zip_ref:
                        zip_ref.extractall(folderPath)

        # r=root, d=directories, f = files
        for r, d, f in os.walk(sourcefilespath):
            for file in f:
                if not ".zip" in file:
                    filepathList.append(os.path.join(r, file))
                    filenameList.append(ntpath.basename(os.path.join(r, file)))
                    result.append([ntpath.basename(os.path.join(r, file)), os.path.join(r, file) ])
                    filecount = filecount + 1

        fileDictionary = dict(result)
        for file in filenameList:
            if re.search("^[0-9]{2}-[0-9]{2}",file):
                value = re.search("^[0-9]{2}-[0-9]{2}",file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                self.progressBar.setValue((progresscount/filecount)*100)
            elif re.search("^[A-Z]{1}\+[A-Z]{1}\-",file):
                value = re.search("^[A-Z]{1}\+[A-Z]{1}\-", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                value1 = re.search("^[A-Z]\+[A-Z]-[A-Z]", file)
                subfolder = value1.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder + "\\" + subfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder + "\\" + subfolder)
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder + "\\" + subfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder + "\\" + subfolder])
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("[A-Z]{2}-[0-9]{2}-[0-9]{2}",file):
                value = re.search("[A-Z]{2}-[0-9]{2}-[0-9]{2}", file)
                newfolder = value.group()
                mod_string = re.sub("[A-Z]{2}-", '', newfolder)
                if not os.path.exists(destinationfilespath + "\\" + mod_string):
                    os.makedirs(destinationfilespath + "\\" + mod_string)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + mod_string)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^FAN",file):
                value = re.search("^FAN", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^ABCD",file):
                value = re.search("^ABCD", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^CASA",file):
                value = re.search("^CASA", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^DTS",file):
                value = re.search("^DTS", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^FI",file):
                value = re.search("^FI", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^IT",file):
                value = re.search("^IT", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^IV",file):
                value = re.search("^IV", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^MC",file):
                value = re.search("^MC", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^NT",file):
                value = re.search("^NT", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^MP",file):
                value = re.search("^MP", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^PG",file):
                value = re.search("^PG", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^TT",file):
                value = re.search("^TT", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^ING",file):
                value = re.search("^ING", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^CAN",file):
                value = re.search("^CAN", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("^BDU",file):
                value = re.search("^BDU", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            else:
                newfolder = "Miscellaneous"
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], destinationfilespath + "\\" + newfolder])
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
        self.runButton.setEnabled(True)
        today = date.today()
        pd.DataFrame(excelData).to_excel(destinationfilespath + "\\" +'output_metadata_' + today.strftime("%b-%d-%Y") + '.xlsx', header=False, index=False)
        log("Total file count: " + str(filecount))

    @QtCore.Slot()
    def runbuttonClicked(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        sourcePath = self.sourceLineEdit.text()
        destinationPath = self.destinationLineEdit.text()

        if not (path.exists(sourcePath) ):
            msg.setText("source folder path invalid")
            msg.setWindowTitle("Info")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return
        elif not (path.exists(destinationPath)):
            msg.setText("destination folder path invalid")
            msg.setWindowTitle("Info")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return
        self.runButton.setEnabled(False)
        self.runOrganizer()
        msg.setText("Export Completed!!!!!")
        msg.setWindowTitle("Info")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()


    def sourcebuttonClicked(self):
        sourcefilespath = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.sourceLineEdit.setText(sourcefilespath)

    def destinationbuttonClicked(self):
        destinationfilespath = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.destinationLineEdit.setText(destinationfilespath)


if __name__ == "__main__":
    app = QtWidgets.QApplication([])

    widget = OrganizerApp()
    widget.resize(800, 200)
    widget.setWindowTitle("Organizer App")
    widget.setWindowIcon(QtGui.QIcon("icon.ico"))
    widget.show()

    sys.exit(app.exec())



