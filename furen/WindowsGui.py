from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog,QProgressBar

class MyWindow(QtWidgets.QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.myButton = QtWidgets.QPushButton(self)
        self.myButton.setObjectName("myButton")
        self.myButton.setText("打开&处理")
        self.myButton.resize(150,50)
        self.pbar=QProgressBar(self)
        self.pbar.setGeometry(0,50,150,20)
        self.myButton.clicked.connect(self.msg)
    def msg(self):
        fileslist, ok1 = QFileDialog.getOpenFileNames(self,"多文件选择","C:/","Excel(*.xls;*.xlsx)")
        print(fileslist, ok1)
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show()
    sys.exit(app.exec_())