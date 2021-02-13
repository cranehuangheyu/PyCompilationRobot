# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import requests #导入requests包
import base64
import math
import xlwt
import json
from bs4 import BeautifulSoup

#设置表格样式
def set_stlye(name, height, bold=False):
    #初始化样式
    style = xlwt.XFStyle()
    #创建字体
    font = xlwt.Font()
    font.bold = bold
    font.colour_index = 0
    font.height = height
    font.name = name
    style.font = font
    return style

#写入excel的实例
def write_excel():
    f = xlwt.Workbook()
    #创建sheet1
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'业务', u'状态', u'北京', u'上海', u'广州', u'深圳', u'状态小计', u'合计']
    column0 = [u'机票', u'船票', u'火车票', u'汽车票', u'其他']
    status = [u'预定', u'出票', u'退票', u'业务小计']

    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i], set_stlye("Time New Roman", 220, True))

    i, j = 1, 0
    while i < 4*len(column0): #控制循环：每次加4
        #第一列
        sheet1.write_merge(i, i+3, 0, 0, column0[j], set_stlye('Arial', 220, True))
        #最后一列
        sheet1.write_merge(i, i+3, 7, 7)
        i += 4
    sheet1.write_merge(21, 21, 0, 1, u'合计', set_stlye("Time New Roman", 220, True))

    i, j = 0, 0
    while i < 4*len(column0): #控制外层循环：每次加4
        for j in range(0, len(status)): #控制内层循环：设置每一行内容
            sheet1.write(i+j+1, 1, status[j])
        i += 4

    #创建sheet2
    sheet2 = f.add_sheet(u'sheet2', cell_overwrite_ok=True)
    row0 = [u'姓名', u'年龄', u'出生日期', u'爱好', u'关系']
    column0 = [u'UZI', u'Faker', u'大司马', u'PDD', u'冯提莫']

    #生成第一行
    for i in range(0, len(row0)):
        sheet2.write(0, i, row0[i], set_stlye('Times New Roman', 220, True))

    #生成第一列
    for i in range(0, len(column0)):
        sheet2.write(i+1, 0, column0[i], set_stlye('Times New Roman', 220, True))

    f.save('data.xls')

def print_hi(name):
    f = xlwt.Workbook()
    # 创建sheet1
    sheet1 = f.add_sheet(u'单选题', cell_overwrite_ok=True)
    sheet2 = f.add_sheet(u'多选题', cell_overwrite_ok=True)
    sheet3 = f.add_sheet(u'判断题', cell_overwrite_ok=True)

    # Use a breakpoint in the code line below to debug your script.
    # print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.
    # url = 'http://www.cntour.cn/'
    get_data_url = 'https://www.zaixian100f.com/exam/test_paper/item/term_id/83.html'
    cookie = ''
    cookieTemp = getCookie(cookie)
    strhtml = requests.get(get_data_url, cookies=cookieTemp)
    soup = BeautifulSoup(strhtml.text, 'lxml')
    # div:nth-child(1)
    dataBase = soup.select('body > div > div.body-wrapper > div.body.paper > div')
    headSoup = soup.select('body > div > div.header-wrapper > div.exam-name.ellipsis')
    xlStyle = set_stlye('Times New Roman', 220, False)
    row1 = 0
    row2 = 0
    row3 = 0
    for item in dataBase:
        data1 = item.contents[3].contents[1]

        # 转换bse64
        temp64 = base64.b64decode(data1.get('date-answer'))
        temp64 = str(temp64).split('\'')
        temp64.pop()
        temp64.pop(0)
        temp64 = ''.join(temp64)
        answerData = json.loads(temp64)

        if data1.get('data-type') == "3":
            data2 = data1.contents[3].contents[1]
            data3 = data1.contents[3].contents[3]

            dateAnswerTemp = answerData["answer"]
            titleTemp = data1.contents[1].text.strip().replace("\n", "")
            option1 = data2.contents[3].contents[3].string.strip().replace("\n", "")
            option2 = data3.contents[3].contents[3].string.strip().replace("\n", "")

            sheet3.write(row3, 0, titleTemp, xlStyle)
            # sheet3.write(row3, 1, result["option1"], xlStyle)
            # sheet3.write(row3, 2, result["option2"], xlStyle)
            sheet3.write(row3, 1, dateAnswerTemp, xlStyle)
            row3 += 1
        elif data1.get('data-type') == "1":
            data2 = data1.contents[3].contents
            conLen = math.ceil((len(data2) - 1) / 2)

            for index in range(0, conLen):
                data3 = data2[index * 2 + 1]
                optionNumTemp = data3.contents[3].contents[3].text.strip().replace("\n", "")
                sheet1.write(row1, 1, optionNumTemp, xlStyle)
                row1 += 1

            dateAnswerTemp = answerData["answer"]
            titleTemp = data1.contents[1].text.strip().replace("\n", "")
            sheet1.write_merge(row1 - conLen, row1 - 1, 0, 0, titleTemp, xlStyle)
            sheet1.write_merge(row1 - conLen, row1 - 1, 2, 2, dateAnswerTemp, xlStyle)
        elif data1.get('data-type') == "2":
            data2 = data1.contents[3].contents
            conLen = math.ceil((len(data2) - 1) / 2)

            for index in range(0, conLen):
                data3 = data2[index * 2 + 1]
                optionNumTemp = data3.contents[3].contents[3].text.strip().replace("\n", "")
                sheet2.write(row2, 1, optionNumTemp, xlStyle)
                row2 += 1

            dateAnswerTemp = answerData["answer"]
            titleTemp = data1.contents[1].string.strip().replace("\n", "")
            sheet2.write_merge(row2 - conLen, row2 - 1, 0, 0, titleTemp, xlStyle)
            sheet2.write_merge(row2 - conLen, row2 - 1, 2, 2, dateAnswerTemp, xlStyle)

    f.save(headSoup[0].string + '.xls')

def getCookie(str):
    cookies = {}  # 初始化cookies字典变量
    for line in str.split(';'):  # 按照字符：进行划分读取
        # 其设置为1就会把字符串拆分成2份
        name, value = line.strip().split('=', 1)
        cookies[name] = value  # 为字典cookies添加内容
    return cookies

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    # write_excel()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
