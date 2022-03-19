import sys
import os
import ntpath
import shutil
import re
from os import path
import pandas as pd
from pyunpack import Archive
from PySide6 import QtCore, QtWidgets, QtGui
from datetime import date
from PySide6.QtWidgets import  QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QLineEdit, QMessageBox, QFileDialog, QProgressBar


DEBUG = True

def log(string):
    if DEBUG:
        print(string)


class FileCountWorker(QtCore.QObject):
    finishedcount = QtCore.Signal()
    progresscount = QtCore.Signal(int)
    reportCountWorker = QtCore.Signal(int)
    sourcefilespath = ""
    destinationfilespath = ""
    def __init__(self, source,destination):
        super().__init__()
        self.sourcefilespath = source
        self.destinationfilespath = destination

    def run(self):
        """Long-running task."""
        filecount = 0
        filepathList = []
        filenameList = []
        excelData = []
        excelData.append(["FileName", "File Path"])
        self.progresscount.emit(20)
        # loop for unzipping files
        for r, d, f in os.walk(self.sourcefilespath):
            for file in f:
                filepathList.append(os.path.join(r, file))
                filenameList.append(ntpath.basename(os.path.join(r, file)))
                excelData.append([ntpath.basename(os.path.join(r, file)), os.path.join(r, file)])
                filecount = filecount + 1
        today = date.today()
        self.progresscount.emit(50)
        excelData.append(["File Count: ", filecount])
        pd.DataFrame(excelData).to_excel(
            self.destinationfilespath + "\\" + 'filesRecieved_List_' + today.strftime("%b-%d-%Y") + '.xlsx', header=False,
            index=False)
        self.progresscount.emit(100)
        log("Total file count: " + str(filecount))
        self.reportCountWorker.emit(filecount)
        self.finishedcount.emit()


class Worker(QtCore.QObject):
    finished = QtCore.Signal()
    progress = QtCore.Signal(int)
    sourcefilespath = ""
    destinationfilespath = ""
    def __init__(self, source, destination):
        super().__init__()
        self.sourcefilespath = source
        self.destinationfilespath = destination

    def run(self):
        """Long-running task."""
        filepathList = []
        filenameList = []
        result = []
        excelData = []

        excelData.append(["Source", "Destination"])

        filecount = 0
        progresscount = 0

        # loop for unzipping files
        for r, d, f in os.walk(self.sourcefilespath):
            for file in f:
                if ".zip" in file:
                    folderPath = os.path.dirname(os.path.abspath(os.path.join(r, file)))
                    Archive(os.path.join(r, file)).extractall(folderPath)
                elif ".tar" in file:
                    folderPath = os.path.dirname(os.path.abspath(os.path.join(r, file)))
                    Archive(os.path.join(r, file)).extractall(folderPath)
                elif ".rar" in file:
                    folderPath = os.path.dirname(os.path.abspath(os.path.join(r, file)))
                    Archive(os.path.join(r, file)).extractall(folderPath)
                elif ".7z" in file:
                    folderPath = os.path.dirname(os.path.abspath(os.path.join(r, file)))
                    Archive(os.path.join(r, file)).extractall(folderPath)


        # r=root, d=directories, f = files
        for r, d, f in os.walk(self.sourcefilespath):
            for file in f:
                if not (".zip" or ".7z" or ".tar" or ".rar") in file:
                    filepathList.append(os.path.join(r, file))
                    filenameList.append(ntpath.basename(os.path.join(r, file)))
                    result.append([ntpath.basename(os.path.join(r, file)), os.path.join(r, file)])
                    filecount = filecount + 1

        fileDictionary = dict(result)
        for file in filenameList:
            if re.search("^[0-9]{2}-[0-9]{2}", file):
                value = re.search("^[0-9]{2}-[0-9]{2}", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^[A-Z]{1}\+[A-Z]{1}\-", file):
                value = re.search("^[A-Z]{1}\+[A-Z]{1}\-", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                value1 = re.search("^[A-Z]\+[A-Z]-[A-Z]", file)
                subfolder = value1.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder + "\\" + subfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder + "\\" + subfolder)
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder + "\\" + subfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder + "\\" + subfolder])
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("[A-Z]{2}-[0-9]{2}-[0-9]{2}", file):
                value = re.search("[A-Z]{2}-[0-9]{2}-[0-9]{2}", file)
                newfolder = value.group()
                mod_string = re.sub("[A-Z]{2}-", '', newfolder)
                if not os.path.exists(self.destinationfilespath + "\\" + mod_string):
                    os.makedirs(self.destinationfilespath + "\\" + mod_string)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + mod_string)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^FAN", file):
                value = re.search("^FAN", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^ABCD", file):
                value = re.search("^ABCD", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^CASA", file):
                value = re.search("^CASA", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^DTS", file):
                value = re.search("^DTS", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^FI", file):
                value = re.search("^FI", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^IT", file):
                value = re.search("^IT", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^IV", file):
                value = re.search("^IV", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^MC", file):
                value = re.search("^MC", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^NT", file):
                value = re.search("^NT", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^MP", file):
                value = re.search("^MP", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^PG", file):
                value = re.search("^PG", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^TT", file):
                value = re.search("^TT", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^ING", file):
                value = re.search("^ING", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^CAN", file):
                value = re.search("^CAN", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            elif re.search("^BDU", file):
                value = re.search("^BDU", file)
                newfolder = value.group()
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
            else:
                newfolder = "Miscellaneous"
                if not os.path.exists(self.destinationfilespath + "\\" + newfolder):
                    os.makedirs(self.destinationfilespath + "\\" + newfolder)
                shutil.copy2(fileDictionary[file], self.destinationfilespath + "\\" + newfolder)
                excelData.append([fileDictionary[file], self.destinationfilespath + "\\" + newfolder])
                progresscount = progresscount + 1
                self.progress.emit((progresscount / filecount) * 100)
        today = date.today()
        self.progress.emit(99)
        pd.DataFrame(excelData).to_excel(
            self.destinationfilespath + "\\" + 'copyifo_metadata_' + today.strftime("%b-%d-%Y") + '.xlsx', header=False,
            index=False)
        self.progress.emit(100)
        log("Total file count: " + str(filecount))
        self.finished.emit()

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
        self.runButton = QPushButton("Run Organizer")
        self.filecountButton = QPushButton("File Count")
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
        self.hLayout3.addWidget(self.filecountButton)
        self.hLayout4.addWidget(self.progressBar)

        self.vLayout.addLayout(self.hLayout1)
        self.vLayout.addLayout(self.hLayout2)
        self.vLayout.addLayout(self.hLayout3)
        self.vLayout.addLayout(self.hLayout4)

        self.runButton.clicked.connect(self.runbuttonClicked)
        self.filecountButton.clicked.connect(self.runFileCount)
        self.sourceButton.clicked.connect(self.sourcebuttonClicked)
        self.destinationButton.clicked.connect(self.destinationbuttonClicked)

    @QtCore.Slot()
    def runOrganizer(self):
        print("running organizer")
        sourcefilespath = self.sourceLineEdit.text()
        destinationfilespath = self.destinationLineEdit.text()
        self.thread = QtCore.QThread()
        self.worker = Worker(sourcefilespath,destinationfilespath)

        # Step 4: Move worker to the thread
        self.worker.moveToThread(self.thread)
        # Step 5: Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.reportProgress)
        # Step 6: Start the thread
        self.thread.start()

        # Final resets
        self.runButton.setEnabled(False)
        self.filecountButton.setEnabled(False)
        self.worker.finished.connect(self.reportFinished)

    @QtCore.Slot()
    def runFileCount(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        self.progressBar.setValue(0)

        sourcefilespath = self.sourceLineEdit.text()
        destinationfilespath = self.destinationLineEdit.text()

        if not (path.exists(sourcefilespath)):
            msg.setText("source folder path invalid")
            msg.setWindowTitle("Info")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return
        elif not (path.exists(destinationfilespath)):
            msg.setText("destination folder path invalid")
            msg.setWindowTitle("Info")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return
        print("running file count")

        self.threadcount = QtCore.QThread()
        self.workercount = FileCountWorker(sourcefilespath, destinationfilespath)

        # Step 4: Move worker to the thread
        self.workercount.moveToThread(self.threadcount)
        # Step 5: Connect signals and slots
        self.threadcount.started.connect(self.workercount.run)
        self.workercount.finishedcount.connect(self.threadcount.quit)
        self.workercount.finishedcount.connect(self.workercount.deleteLater)
        self.threadcount.finished.connect(self.threadcount.deleteLater)
        self.workercount.progresscount.connect(self.reportProgress)
        self.workercount.reportCountWorker.connect(self.reportCount)
        # Step 6: Start the thread
        self.threadcount.start()

        # Final resets
        self.filecountButton.setEnabled(False)
        self.runButton.setEnabled(False)
        self.workercount.finishedcount.connect(self.reportFinishedCount)



    @QtCore.Slot()
    def runbuttonClicked(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        self.progressBar.setValue(0)

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

    @QtCore.Slot()
    def reportProgress(self, value):
        self.progressBar.setValue(value)

    @QtCore.Slot()
    def reportFinished(self):
        self.runButton.setEnabled(True)
        self.filecountButton.setEnabled(True)
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Export completed successfully")
        msg.setWindowTitle("Info")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()

    @QtCore.Slot()
    def reportFinishedCount(self):
        self.runButton.setEnabled(True)
        self.filecountButton.setEnabled(True)

    @QtCore.Slot()
    def reportCount(self, filecount):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Files count :" + str(filecount))
        msg.setInformativeText("File Names exported:" + self.destinationLineEdit.text())
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



