import sys

from excel import *

from PyQt5.QtCore import QLine, QObject
from PyQt5.QtWidgets import (QMainWindow, QTextEdit,
                             QAction, QFileDialog, QApplication, QLabel, QLineEdit, QGridLayout, QPushButton, QComboBox,
                             QVBoxLayout, QHBoxLayout, QWidget, QDialog, QMessageBox)
from PyQt5.QtGui import QIcon


class Application(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()


    def initUI(self):


        chFileBtn = QPushButton("Выбор файла .xlxs", self)
        colNum = QLabel("Колонка в Excel с номерами:")
        listName = QLabel('Имя листа в Excel файле:')
        startBtn=QPushButton("Запуск программы")

        self.chFileField = QLineEdit(self)
        self.chFileField.setReadOnly(True)
        self.listNameEdit = QLineEdit(self)
        self.listNameEdit.setText("Лист1")


        colNumBox = QComboBox(self)
        colNumBox.addItem(None)
        colNumBox.addItem("1")
        colNumBox.addItem("2")
        colNumBox.addItem("3")
        colNumBox.addItem("4")
        colNumBox.addItem("5")
        colNumBox.activated[str].connect(self.onActivated)


        chFileBtn.clicked.connect(self.showDialog)
        startBtn.clicked.connect(self.startProgramm)


        hbox= QHBoxLayout()
        hbox.addWidget(colNum)
        hbox.addWidget(colNumBox)
        hbox.addStretch()

        hbox2=QHBoxLayout()
        hbox2.addWidget(listName)
        hbox2.addWidget(self.listNameEdit)

        vbox = QVBoxLayout()
        vbox.addWidget(chFileBtn)

        vbox.addWidget(self.chFileField)

        vbox.addLayout(hbox)

        vbox.addLayout(hbox2)

        vbox.addWidget(startBtn)




        self.setLayout(vbox)

        self.setGeometry(500, 500, 400, 200)
        self.setWindowTitle('Поиск копий')
        self.setWindowIcon(QIcon('icon.png'))
        self.show()


    def showDialog(self):

        fname = QFileDialog.getOpenFileName(self, 'Выбор файла', 'c:\\')
        print(fname)
        array=str(fname)
        array=array.split(',')
        print(array[0])
        string=array[0]
        string = string.replace("/", "\\")
        string = string.replace("\'", "")
        string = string.replace("(", "")
        string = string.replace("/", "\\")
        print(string)
        self.path=string
        self.chFileField.setText(string)

    def onActivated(self, text):
        if len(text)==1:
            self.colNubValue=int(text)
            print(self.colNubValue)
        else:
            print("try again!")

    def onChanged(self, text):
        self.listValue.setText(text)

    def dialog(self):

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowIcon(QIcon('error.png'))

        msg.setText("Что-то пошло не так...")
        msg.setInformativeText("Проверьте, что все парамеры выбраны")
        msg.setWindowTitle("Ошибка")
        msg.setDetailedText("Скорее всего какие-то параметры были введены неправильно либо не были заданы."
                            "Проверьте правильность введённых даннхы и повторите:")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        retval = msg.exec_()

    def ok(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('ok.png'))
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText("Программа завершена успешно!")
        msgBox.setWindowTitle("Выполнено")
        msgBox.setStandardButtons(QMessageBox.Ok)

        returnValue = msgBox.exec()
        if returnValue == QMessageBox.Ok:
            print('OK clicked')

    def startProgramm(self):
        try:
            file_name=self.path
            list_name=self.listNameEdit.text()
            column=self.colNubValue
            print(list_name)
            print("create obj")
            if file_name!=None or list_name!=None or column!=None:
                obj = ExcelOpen()
                obj.open_file(file_name,list_name,column)

                obj.read_and_write()
                obj.save_file()
            else:
                self.dialog()
        except Exception:
            self.dialog()
            return None
        self.ok()







if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Application()
    sys.exit(app.exec_())