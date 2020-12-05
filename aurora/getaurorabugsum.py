import xlwt,zipfile,os,shutil
from lxml import etree
from pyfiglet import Figlet
#请将该脚本放到需要读取的zip文件同目录下
#该脚本会自动解压该目录下所有zip文件并读取绿盟的index.heml文件，提取高、中危漏洞总数并写入表格中
#执行完成会自动删除解压的所有文件

auroranames=[] #极光自动化文件名缓存
auroraList=[]  #漏洞数量列表缓存
def checkfile():   
    files = os.listdir('.') #遍历当前文件夹下的所有文件
    for afile in files:
        ext=afile.split('_')[-1] 
        #print(ext)
        if(ext=='html.zip'):# 判定 zip文件 报告内容 
            aurora=unzip(afile)#解压文件到当前目录
            auroranames.append(aurora)
    
def unzip(filename):
    zfile=zipfile.ZipFile(filename,'r')
    for afile in zfile.namelist():
        zfile.extract(afile,'./已统计/'+filename[:-4])
    files=os.listdir('./已统计/'+filename[:-4])
    zfile.close()
    return filename[:-4]

def getaurora(getname):
    htmlname='./已统计/'+getname+'/index.html'
    htmltext=open(htmlname,'r',encoding='utf-8').read()
    html=etree.HTML(htmltext)   #转换格式
    name=html.xpath('//*[@id="content"]/div[2]/div[2]/table[2]//td[1]//tr[1]/td/text()')#任务名称
    high=html.xpath('//*[@id="content"]/div[6]/div[2]//tfoot//td[2]/span/text()')#高危漏洞数量
    medium=html.xpath('//*[@id="content"]/div[6]/div[2]//tfoot//td[3]/text()')#中危险漏洞数
    temp=name+high+medium   #合并数据
    auroraList.append(temp)       #形成列表
    
def bugnum(outfilename): #生成xls模板
    new_workbook = xlwt.Workbook()                              #新建工作簿
    worksheet = new_workbook.add_sheet('网管中心')               #新建工作表
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
    for a in range(0,len(title)):
        worksheet.write(1,a,title[a],style0)           #往工作表中的指定单元格写入内容,并指定style0格式
    for i in range(0,len(auroraList)):  #循环列表写入数据
        worksheet.write(i+2, 0, auroraList[i][0], style1)
        worksheet.write(i+2, 1, auroraList[i][1], style1)
        worksheet.write(i+2, 2, auroraList[i][2], style1) 
        print('已写入'+str(i+1)+'行数据')
    new_workbook.save('./绿盟高、中危漏洞统计.xls')



if __name__ == '__main__':
    checkfile()
    for name in auroranames:
        getaurora(name)
    shutil.rmtree('./已统计/') # 移除相关文件目录
    bugnum(auroraList)
    print('已删除解压数据')
    f = Figlet(font='slant')
    print(f.renderText('completed'))
