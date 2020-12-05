
import hashlib, hmac, json, os, sys,re,time,requests,xlwt,zipfile,shutil,getopt
from datetime import datetime
from lxml import etree
data=[]
result=[]
ports=[]
def getnessus(initname): #得到nessus html扫描报告 模板数据
    htmltext=open(initname,'r',encoding='utf-8').read()
    html=etree.HTML(htmltext)
    vultype=html.xpath('//*[@id="report"]/div[3]/ul/li/a/text()')[0]
    if 'Vulnerabilities by Plugin' not in vultype:
        print('Nessus 报告模板非plugin模板，请检查文件模板')
        return
    idname=html.xpath('//*[@id="report"]/div[3]/ul/li/ul/li/a/@href')
    textname=html.xpath('//*[@id="report"]/div[3]/ul/li/ul/li/a/text()')
    id=''#获取id
    for i in range(len(textname)):
        if 'SYN scanner' in textname[i]:
            id=(idname[i]+'-container').replace('#','')
    msg=html.xpath('//*[@id="'+id+'"]/h2/text()')
    for m in msg:
        ip=re.findall( r'[0-9]+(?:\.[0-9]+){3}',m)[0]
        port=re.findall(r'(?:/)\d+',m)[0][1:]
        data.append([ip,port])
    for node in data:
        if node[1] not in ports:
            ports.append(node[1])
    for port in ports:
        ipstr=''
        for node in data:
            if port==node[1]:
                ipstr=ipstr+node[0]+'\n'
        ipstr=ipstr[:-1]
        result.append([port,ipstr,str(len(ipstr.split()))])
def savefile(outfilename,dataval):
    wb=xlwt.Workbook()
    ws=wb.add_sheet('SYN信息')
    # title设置
    titlestyle = xlwt.XFStyle()
    # 设置字体
    titlefont = xlwt.Font()
    titlefont.name='SimSun'
    titlefont.height=20*11
    titlestyle.font = titlefont
    # 标题单元格对齐方式
    titlealignment = xlwt.Alignment()
    # 水平对齐方式和垂直对齐方式
    titlealignment.horz = xlwt.Alignment.HORZ_CENTER
    titlealignment.vert = xlwt.Alignment.VERT_CENTER
    # 自动换行
    titlealignment.wrap = 1
    titlestyle.alignment = titlealignment
    # 单元格背景设置
    titlepattern = xlwt.Pattern()
    titlepattern.pattern = xlwt.Pattern.SOLID_PATTERN
    titlepattern.pattern_fore_colour = xlwt.Style.colour_map['sky_blue'] # 设置单元格背景颜色为蓝
    titlestyle.pattern = titlepattern
    # 单元格边框
    titileborders = xlwt.Borders()
    titileborders.left = 1
    titileborders.right = 1
    titileborders.top = 1
    titileborders.bottom = 1
    titileborders.left_colour = 0x40
    titlestyle.borders=titileborders
    # 设置标题
    ws.write(0, 0, '序号',titlestyle)
    ws.write(0, 1, '端口号',titlestyle)
    ws.write(0, 2, 'IP地址',titlestyle)
    ws.write(0, 3, '数量',titlestyle)
    ws.write(0, 4, '动作',titlestyle)
    ws.write(0, 5, '备注',titlestyle)
    # 设置列高度
    ws.row(0).height_mismatch = True
    ws.row(0).height= int(20 * 40 )
    # 设置列宽度
    ws.col(0).width = int(256 * 8)
    ws.col(1).width = int(256 * 10)
    ws.col(2).width = int(256 * 46)
    ws.col(3).width = int(256 * 17)
    ws.col(4).width = int(256 * 17)
    ws.col(5).width = int(256 * 17)


    contentstyle=xlwt.XFStyle()
    contentalignment = xlwt.Alignment()
    contentalignment.wrap = 1
    # 水平对齐方式和垂直对齐方式
    contentalignment.horz = xlwt.Alignment.HORZ_CENTER
    contentalignment.vert = xlwt.Alignment.VERT_CENTER
    contentstyle.alignment=contentalignment

    # 写入数据
    for i in range(len(result)):
        serials=i+1
        ws.write(serials, 0, str(serials),contentstyle)
        ws.write(serials, 1, dataval[i][0],contentstyle)
        ws.write(serials, 2, dataval[i][1],contentstyle)
        ws.write(serials, 3, dataval[i][2],contentstyle)
        ws.write(serials, 4, '开启',contentstyle)
    # 保存excel文件
    wb.save(outfilename)
if __name__ == "__main__":
    arg=sys.argv[1]
    getnessus(arg)
    savefile(arg[:-5].replace('plugin','')+'SYN信息.xls',result)