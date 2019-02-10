# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import PyQt5
from PyQt5 import QtCore, QtWidgets

class MyWindow(QtWidgets.QWidget):
    def __init__(self, parent = None):
        QtWidgets.QWidget.__init__(self, parent)
        self.label = QtWidgets.QLabel("Создаем раскрывающийся список")
        self.label.setAlignment(QtCore.Qt.AlignHCenter)
        self.cb = QtWidgets.QCombobox()
        self.cb.addItems(1, 2, 3, 4, 5, 6, 'uk')
        self.btnQuit = QtWidgets.QPushButton("Выбрать значеие")
        self.vbox = QtWidgets.QVBoxLayout()
        self.vbox.addWidget(self.label)
        self.vbox.addWidget(self.cb)
        self.vbox.addWidget(self.btnQuit)
        self.setLayout(self.vbox)
        self.btnQuit.clicked.connect(QtWidgets.qApp.quit)

if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindow()
    window.setWindowTitle("ООП - стиль создания окна")
    window.resize(300, 70)
    window.show()
    sys.exit(app.exec_())
    
