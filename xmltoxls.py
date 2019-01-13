import xlwt
import xml.dom.minidom
import re

inputpath = 'f:\\Projects\\ordinaryPy\\input.xml'
outputpath = 'f:\\Projects\\ordinaryPy\\output.xls'
#获取父节点
def getparen(i,testcases):
    c = testcases[i].parentNode
    s = ""
    while c.nodeName == "testsuite":
        s = c.getAttribute("name")+" / "+s
        c = c.parentNode
    s = "T"+str(i+1)+"、"+s+testcases[i].getAttribute("name")
    return s

#正则表达式替换函数
def replace(str):
        s1 = re.sub(r'<.*?>','',str)
        s2 = re.sub(r'&ldquo;','“',s1)
        s3 = re.sub(r'&rdquo;','”',s2)
        s4 = re.sub(r'&.*?;','',s3)
        return s4

#行数计数器
j = 0
#初始化样式
style = xlwt.XFStyle() # 初始化样式
style.alignment.horz = xlwt.Alignment.HORZ_CENTER # 垂直对齐
style.alignment.vert = xlwt.Alignment.VERT_CENTER # 水平对齐
style.alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT # 自动换行
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('NewSheet',cell_overwrite_ok=True)
row0 = ['testcase','summary','preconditions','stepnumber','ations','expectedresults']
for i in range(0,5):
    worksheet.write(0,i,row0[i],style)
#解析XML文件
root = xml.dom.minidom.parse(inputpath)
testcases = root.getElementsByTagName('testcase')
#写入excel
for i in range(testcases.length):
    worksheet.col(0).width = 8000
    worksheet.col(1).width = 10000
    worksheet.col(2).width = 6000
    worksheet.col(3).width = 3333
    worksheet.col(4).width = 10000
    worksheet.col(5).width = 10000
    worksheet.write(j+1,0,getparen(i,testcases),style) 
    if testcases[i].getElementsByTagName('summary')[0].childNodes.length>0:
        #利用正则表达式替换多余的标签
        temp = replace(testcases[i].getElementsByTagName('summary')[0].childNodes[0].data)
        worksheet.write(j+1,1,temp,style)
    if testcases[i].getElementsByTagName('preconditions')[0].childNodes.length>0:
        #利用正则表达式替换多余的标签
        temp = replace(testcases[i].getElementsByTagName('preconditions')[0].childNodes[0].data)
        worksheet.write(j+1,2,temp,style)
    stepnumbers = testcases[i].getElementsByTagName("step_number")
    actions = testcases[i].getElementsByTagName("actions")
    expected = testcases[i].getElementsByTagName("expectedresults")
    for k in range(stepnumbers.length):
        j=j+1
        worksheet.write(j,3,stepnumbers[k].childNodes[0].data,style)
        if actions[k].childNodes.length>0:
                #利用正则表达式替换多余的标签
                temp = replace(actions[k].childNodes[0].data)
                worksheet.write(j,4,temp,style)
        #判断子节点是否存在
        if expected[k].childNodes.length>0:
                #替换正则表达式
                temp = replace(expected[k].childNodes[0].data)
                worksheet.write(j,5,temp,style)
workbook.save(outputpath)