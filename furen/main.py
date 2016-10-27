import os,re
import win32com.client

#清理excel进程
os.system('taskkill /F /IM excel.exe')
excel=win32com.client.Dispatch('Excel.Application')
filelist=os.listdir('files')
l = [[] for row in range(len(filelist))]
temp=0
for file in filelist:
    AB=0
    C=0
    xlbook = excel.Workbooks.Open('c:\\furen\\files\\'+file)
    sht = xlbook.Worksheets('表格名称')
    #姓名
    name=re.findall(r'姓名：(\D*)职位',sht.Cells(5,1).value)[0].replace(' ', '')
    print(name)
    for i in range (4,35):
        try:
            if sht.Cells(i,1).value=='总计AB类求和':
                for r in range(2,10):
                    if sht.Cells(i,r).value!=None:
                        AB = re.findall(r'\d*\.*\d*', str(sht.Cells(i, r).value))[0]
                        print('AB:'+str(AB))
                        break
        except:
            pass
        try:
            if sht.Cells(i, 1).value == '总计C类求和':
                for r in range(2, 10):
                    if sht.Cells(i, r).value != None:
                        C=re.findall(r'\d*\.*\d*',str(sht.Cells(i, r).value))[0]
                        print('C:' + str(C))
                        break
        except:
            pass
    xlbook.Close(SaveChanges=0)
    l[temp].append(name)
    l[temp].append(AB)
    l[temp].append(C)
    temp=temp+1
print(l)
print('共计合并'+str(len(filelist))+'个文件，导出'+str(len(l))+'条数据。')
excel.quit()
#写入新excel
excel=win32com.client.Dispatch('Excel.Application')
excel.Visible = 1
xlsBook = excel.Workbooks.Add()
xlsSht = xlsBook.Worksheets('Sheet1')
xlsSht.Cells(1,1).Value='姓名'
xlsSht.Cells(1,2).Value='AB类'
xlsSht.Cells(1,3).Value='C类'
for i in range(0,len(l)):
    xlsSht.Cells(i+2, 1).Value = l[i][0]
    xlsSht.Cells(i+2, 2).Value = l[i][1]
    xlsSht.Cells(i+2, 3).Value = l[i][2]