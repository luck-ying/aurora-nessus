import xlwt
import os
import csv
from pyfiglet import Figlet  #字体
#请将该脚本放到需要读取的cvs文件同目录下csv
#该脚本会自动d读取该目录下所有c's'v文件并统计严重、高、中危漏洞总数并写入表格中

files = os.listdir('.') #遍历当前文件夹下的所有文件
column_data=[]
nessusList=[]
critical_num=set()#定义集合
high_num=set()
medium_num=set()

def getnessuscsv():#依次读取cvs
    for afile in files:
        ext=afile.split('.')[-1]
        if(ext=='csv'):# 判定 csv文件 
            with open(afile, 'r',encoding='utf-8') as csvfile:    
                reader  = csv.reader(csvfile)   
                column= [row[3]for row in reader]
                for a in column:     
                    if a=='Critical':
                        critical_num.add(a)
                    if a=='High':
                        high_num.add(a)
                    if a=='Medium':
                        medium_num.add(a)
                nessus=afile.split('.')[0],column.count('Critical'),column.count('High'),column.count('Medium')
                nessusList.append(nessus)
            csvfile.close() #关闭表格
    return nessusList

def savefile():
    new_workbook = xlwt.Workbook()                              #新建工作簿
    worksheet = new_workbook.add_sheet('sheet')               #新建工作表
    style0 = xlwt.XFStyle()                                     #初始化一个样式
    pattern = xlwt.Pattern()                                    #初始化一个单元格的背景颜色
    font0 = xlwt.Font()                                         #初始化一个字体格式
    font0.name = '等线'
    font0.bold = True                                           #加粗
    font0.height = 240                                          #字体大小，240/20=12
    style0.font = font0                                         #将初始化的字体格式font0赋给style.font
    borders0 = xlwt.Borders()                                   #初始化一个边框格式
    borders0.top = xlwt.Borders.THIN
    borders0.bottom = xlwt.Borders.THIN
    borders0.left = xlwt.Borders.THIN
    borders0.right = xlwt.Borders.THIN
    style0.borders = borders0                                   #将初始化的边框格式borders0赋给style.borders
    alignment0 = xlwt.Alignment()                               #初始化对齐格式
    alignment0.horz = xlwt.Alignment.HORZ_CENTER                # 左右居中
    alignment0.vert = xlwt.Alignment.VERT_CENTER                # 上下居中
    style0.alignment = alignment0                               #将初始化的边框格式borders0赋给style.borders
    style1 = xlwt.XFStyle()
    font1 = xlwt.Font()
    font1.name = '等线'
    font1.bold = False
    font1.height = 220  # 字体大小乘以20
    style1.font = font1
    borders1 = xlwt.Borders()
    borders1.top = xlwt.Borders.THIN
    borders1.bottom = xlwt.Borders.THIN
    borders1.left = xlwt.Borders.THIN
    borders1.right = xlwt.Borders.THIN
    style1.borders = borders1
    alignment1 = xlwt.Alignment()
    alignment1.horz = xlwt.Alignment.HORZ_LEFT
    alignment1.vert = xlwt.Alignment.VERT_CENTER
    style1.alignment = alignment1
    worksheet.write_merge(0,0,0,0,'系统名称',style0)           #标题:系统名称
    worksheet.write_merge(0,0,1,3,'绿盟',style0)                #标题:绿盟
    worksheet.write_merge(0,0,4,7, 'nessus',style0)             #标题:Nessus
    title=['  ','高危','中危','合计','严重','高危','中危','合计']    #第二行
    for t in range(0,len(title)):
        worksheet.write(1,t,title[t],style0)           #往工作表中的指定单元格写入内容,并指定style0格式
    for i in range(0,len(nessusList)):  #循环列表写入数据
        worksheet.write(i+2, 0, nessusList[i][0], style1) 
        worksheet.write(i+2, 4, nessusList[i][1], style1)
        worksheet.write(i+2, 5, nessusList[i][2], style1)
        worksheet.write(i+2, 6, nessusList[i][3], style1) 
        print('已写入'+str(i+1)+'行数据')  
    new_workbook.save('./nessus严重、高、中危漏洞统计.xls')



if __name__ == '__main__':
    getnessuscsv()
    savefile()
    f = Figlet(font='slant')
    print(f.renderText('completed'))




