import http.cookiejar
from urllib.parse import urlencode
from urllib.request import build_opener, HTTPCookieProcessor, Request
import xlwt
from PIL import Image
import matplotlib.pyplot as plt
from lxml import etree
from bs4 import BeautifulSoup
import xlrd  # 导入xlrd模块
import os

path1 = os.path.abspath('.')  # 表示当前所处的文件夹的绝对路径
cookie = http.cookiejar.CookieJar()
opener = build_opener(HTTPCookieProcessor(cookie))
URL = 'http://URP'                   #将 URP 替换成本校的教务处地址
loginURL = URL + '/loginAction.do'  # URP教务处登陆地址
srcImage = URL + '/validateCodeAction.do'  # 验证码图片
gradeURL = URL + '/gradeLnAllAction.do?type=ln&oper=qbinfo'  # 登陆后的实际成绩显示页面
formDate = {'zjh': '', 'mm': '', 'v_yzm': ''}


def getAccount():
    formDate['zjh'] = input("请输入你的学号：")
    formDate['mm'] = input("请输入你的密码：")


def getValidationCode():
    picture = opener.open(srcImage).read()
    with open('vimage.jpg', 'wb') as f:
        f.write(picture)
    img = Image.open('vimage.jpg')
    plt.figure("code", figsize=(1.2, 2.4))
    plt.axis('off')
    plt.imshow(img)
    plt.show()
    formDate['v_yzm'] = input('请输入验证码：')


def getGradePage():
    getAccount()
    error = True
    while error:
        if not formDate['v_yzm'] or error:
            getValidationCode()
        postedData = urlencode(formDate).encode(encoding='UTF-8')
        request = Request(loginURL, postedData)
        loginPage = opener.open(request)
        loginTree = etree.HTML(loginPage.read())
        error = loginTree.xpath('//td[@class="errorTop"]')  # 验证码错误、密码错误等错误都会出现此标志
    html = opener.open(gradeURL)
    main = html.read().decode('gbk')
    soup = BeautifulSoup(main, 'lxml')
    content = soup.find_all('td', align="center")
    data_list = []
    for data in content:
        data_list.append(data.text.strip())
    new_list = [data_list[i:i + 7] for i in range(0, len(data_list), 7)]
    book = xlwt.Workbook()
    sheet1 = book.add_sheet('sheet1', cell_overwrite_ok=True)
    heads = ['课程号', '课序号', '课程名', '英文课程名', '学分', '课程属性', '成绩']
    print("准备将数据写入表格...")
    ii = 0
    for head in heads:
        sheet1.write(0, ii, head)
        ii += 1
    i = 1
    for list in new_list:
        j = 0
        for data in list:
            sheet1.write(i, j, data)
            j += 1
        i += 1
    book.save(formDate['zjh']+".xls")
    print('写入成功！')


def readExcel(data_path, sheetname):
    print("读取成绩数据\n准备输出成绩")
    data_path = data_path  # excle表格路径，需传入绝对路径
    sheetname = sheetname  # excle表格内sheet名
    data = xlrd.open_workbook(data_path)  # 打开excel表格
    table = data.sheet_by_name(sheetname)  # 切换到相应sheet
    keys = table.row_values(0)  # 第一行作为key值
    rowNum = table.nrows  # 获取表格行数
    colNum = table.ncols  # 获取表格列数
    if rowNum < 2:
        print("excle内数据行数小于2")
    else:
        L = []  # 列表L存放取出的数据
        for i in range(1, rowNum):  # 从第二行（数据行）开始取数据
            sheet_data = {}  # 定义一个字典用来存放对应数据
            for j in (2, 4, 5, 6):  # j对应列值
                sheet_data[keys[j]] = table.row_values(i)[j]  # 把第i行第j列的值取出赋给第j列的键值，构成字典
            L.append(sheet_data)  # 一行值取完之后（一个字典），追加到L列表中
    # return L
    zjd = 0.
    zxf = 0.
    for kemu in L:
        mc=kemu['课程名']
        cj = kemu['成绩']
        xf = kemu['学分']
        cj1 = flota(cj)
        xf1 = float(xf)
        jd = GPA(cj1)
        zjd += jd * xf1
        zxf = zxf + xf1
        print("课程名 %s 成绩 %s %s 学分 %s 绩点 %s" % (mc, cj, cj1, xf1, '%.1f' %jd))
    jjj= zjd/zxf
    jjd= float('%.2f' % jjj)
    print(jjd,jjj)

'''绩点转换说明（尚没有测试有挂科的情况）
绩点（包括体育和选修)
1. 90-100分折合为4.0绩点，优秀折合4.0绩点  
2. 80-89分折合为3.0-3.9绩点（80分折合为3.0绩点，81分折合为3.1绩点，余者类推，下同），良好折合3.5绩点
3. 70-79分折合为2.0-2.9绩点，中等折合2.5绩点  
4. 60-69分折合为1.0-1.9绩点，及格折合1.5绩点  
5. 60分以下（不及格）折合0绩点  
总绩点=∑(已修读课程绩点×课程学分)/∑已修读课程学分
'''
def GPA(a):
    if a >= 90:
        return float(4)
    elif a < 90 and a >= 60:
        return a / 10. - 5
    elif a < 60:
        return float(0)


def flota(a):
    if a == '优秀':
        return float(95)
    elif a == '良好':
        return float(85)
    elif a == '中等':
        return float(75)
    elif a == '及格':
        return float(65)
    elif a == '不及格':
        return float(0)
    else:
        return float(a)

if URL == 'http://URP':
    URL = input('请输入本校URP教务系统的网址或者IP地址：')
    loginURL = 'http://' + URL + '/loginAction.do'  # URP教务处登陆地址
    srcImage = 'http://' + URL + '/validateCodeAction.do'  # 验证码图片
    gradeURL = 'http://' + URL + '/gradeLnAllAction.do?type=ln&oper=qbinfo'  # 登陆后的实际成绩显示页面
    a = getGradePage()
else:
    a = getGradePage()

b = readExcel(path1 + r'\\' + formDate['zjh']+".xls",'sheet1')
c = input('按回车键退出')
