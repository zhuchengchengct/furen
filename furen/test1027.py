import os,re
import win32com.client

#清理excel进程
os.system('taskkill /F /IM excel.exe')
excel = win32com.client.Dispatch('Excel.Application')
filelist=os.listdir('files')
l = [[] for row in range(len(filelist))]
temp=0
for file in filelist:
    print(file)
    countA=0
    countB=0
    countC=0
    xlbook = excel.Workbooks.Open('c:\\py\\furen\\files\\'+file)
    try:
        sht = xlbook.Worksheets('表格名称')
        #姓名
        name=re.findall(r'姓名：(\D*)职位',sht.Cells(5,1).value)[0].replace(' ', '')
        print(name)
        #加班小时数统计，按照加班日期非空项
        t=[i for i in range(8,40) if sht.Cells(i,2).value==None][0]
        for n in range (8,t):
            if re.findall(r'\w',sht.Cells(n,1).value)[0].upper()=='A':
                countA=countA+float(re.findall(r'\d*\.*\d*',str(sht.Cells(n,8).value))[0])
            elif re.findall(r'\w',sht.Cells(n,1).value)[0].upper()=='B':
                countB = countB + float(re.findall(r'\d*\.*\d*', str(sht.Cells(n, 8).value))[0])
            elif re.findall(r'\w',sht.Cells(n,1).value)[0].upper()=='C':
                countC = countC + float(re.findall(r'\d*\.*\d*', str(sht.Cells(n, 8).value))[0])
        print(countA,countB,countC)
        l[temp].append(name)
        l[temp].append(countA+countB)
        l[temp].append(countC)
        temp=temp+1
    except:
        l[temp].append('格式错误！夫人请检查以下文件：'+file)
        temp=temp+1
excel.quit()
print(l)
print('共计合并' + str(len(filelist)) + '个文件，导出' + str(len(l)) + '条数据。')

#写入新excel
os.system('taskkill /F /IM excel.exe')
excel=win32com.client.Dispatch('Excel.Application')
excel.Visible = 1
xlsBook = excel.Workbooks.Add()
xlsSht = xlsBook.Worksheets('Sheet1')
xlsSht.Cells(1,1).Value='姓名'
xlsSht.Cells(1,2).Value='AB类'
xlsSht.Cells(1,3).Value='C类'
for i in range(0,len(l)):
    try:
        xlsSht.Cells(i+2, 1).Value = l[i][0]
        xlsSht.Cells(i+2, 2).Value = l[i][1]
        xlsSht.Cells(i+2, 3).Value = l[i][2]
    except:
        pass