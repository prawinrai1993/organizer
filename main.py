import sys
import os
import ntpath
import shutil
import re
from os import path
from PySide6 import QtCore, QtWidgets, QtGui
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

    def runOrganizer(self):
        print("running organizer")
        sourcefilespath = self.sourceLineEdit.text()
        destinationfilespath = self.destinationLineEdit.text()
        filepathList = []
        filenameList = []
        result = []

        filecount = 0
        progresscount = 0
        self.progressBar.setValue(0)
        # r=root, d=directories, f = files
        for r, d, f in os.walk(sourcefilespath):
            for file in f:
                filepathList.append(os.path.join(r, file))
                filenameList.append(ntpath.basename(os.path.join(r, file)))
                result.append([ntpath.basename(os.path.join(r, file)), os.path.join(r, file) ])
                filecount = filecount + 1

        log(filenameList)
        fileDictionary = dict(result)
        for file in filenameList:
            if re.search("^[0-9]{2}-[0-9]{2}",file):
                value = re.search("^[0-9]{2}-[0-9]{2}",file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
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
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("[A-Z]{2}-[0-9]{2}-[0-9]{2}",file):
                value = re.search("[A-Z]{2}-[0-9]{2}-[0-9]{2}", file)
                newfolder = value.group()
                mod_string = re.sub("[A-Z]{2}-", '', newfolder)
                if not os.path.exists(destinationfilespath + "\\" + mod_string):
                    os.makedirs(destinationfilespath + "\\" + mod_string)

                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + mod_string)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("FAN",file):
                value = re.search("FAN", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)

                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            elif re.search("ABCD",file):
                value = re.search("ABCD", file)
                newfolder = value.group()
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)

                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)
            else:
                newfolder = "Miscellaneous"
                print(file)
                if not os.path.exists(destinationfilespath + "\\" + newfolder):
                    os.makedirs(destinationfilespath + "\\" + newfolder)
                shutil.copy2(fileDictionary[file], destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progressBar.setValue((progresscount / filecount) * 100)






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
        self.runOrganizer()

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



