from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog,QProgressBar
import os,re
import win32com.client

class MyWindow(QtWidgets.QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.myButton = QtWidgets.QPushButton(self)
        self.myButton.setObjectName("myButton")
        self.myButton.setText("选择需要处理的文件")
        self.myButton.resize(150, 50)
        self.pbar=QProgressBar(self)
        self.pbar.setGeometry(0,50,150,20)
        self.myButton.clicked.connect(self.editfile)
    def editfile(self):
        # 清理excel进程
        os.system('taskkill /F /IM excel.exe')
        filelist, ok1 = QFileDialog.getOpenFileNames(self, "多文件选择", "C:/py/furen/files/", "Excel(*.xls;*.xlsx)")
        if filelist:
            excel = win32com.client.Dispatch('Excel.Application')
            l = [[] for row in range(len(filelist))]
            print(l)
            temp = 0
            for file in filelist:
                print(file)
                countA = 0
                countB = 0
                countC = 0
                xlbook = excel.Workbooks.Open(file)
                try:
                    sht = xlbook.Worksheets('表格名称')
                    # 姓名
                    name = re.findall(r'姓名：(\D*)职位', sht.Cells(5, 1).value)[0].replace(' ', '')
                    print(name)
                    # 加班小时数统计，按照加班日期非空项
                    t = [i for i in range(8, 40) if sht.Cells(i, 2).value == None][0]
                    for n in range(8, t):
                        if re.findall(r'\w', sht.Cells(n, 1).value)[0].upper() == 'A':
                            countA = countA + float(re.findall(r'\d*\.*\d*', str(sht.Cells(n, 8).value))[0])
                        elif re.findall(r'\w', sht.Cells(n, 1).value)[0].upper() == 'B':
                            countB = countB + float(re.findall(r'\d*\.*\d*', str(sht.Cells(n, 8).value))[0])
                        elif re.findall(r'\w', sht.Cells(n, 1).value)[0].upper() == 'C':
                            countC = countC + float(re.findall(r'\d*\.*\d*', str(sht.Cells(n, 8).value))[0])
                    print(countA, countB, countC)
                    l[temp].append(name)
                    l[temp].append(countA)
                    l[temp].append(countB)
                    l[temp].append(countC)
                    temp = temp + 1
                    # 进度条测试
                    self.pbar.setValue(temp / len(filelist)*100)
                except:
                    l[temp].append('格式错误！夫人请检查以下文件：' + file)
                    temp = temp + 1
                    self.pbar.setValue(temp / len(filelist) * 100)
            excel.quit()
            print(l)
            print('共计合并' + str(len(filelist)) + '个文件，导出' + str(len(l)) + '条数据。')
            # 写入新excel
            os.system('taskkill /F /IM excel.exe')
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = 1
            xlsBook = excel.Workbooks.Add()
            xlsSht = xlsBook.Worksheets('Sheet1')
            xlsSht.Cells(1, 1).Value = '姓名'
            xlsSht.Cells(1, 2).Value = 'A类'
            xlsSht.Cells(1, 3).Value = 'B类'
            xlsSht.Cells(1, 4).Value = 'C类'
            for i in range(0, len(l)):
                try:
                    xlsSht.Cells(i + 2, 1).Value = l[i][0]
                    xlsSht.Cells(i + 2, 2).Value = l[i][1]
                    xlsSht.Cells(i + 2, 3).Value = l[i][2]
                    xlsSht.Cells(i + 2, 4).Value = l[i][3]
                except:
                    pass
        else:
            pass
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show()
    sys.exit(app.exec_())