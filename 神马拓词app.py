# coding=utf-8
import os
import pandas as pd
import xlrd
from tkinter import *
from xlutils.copy import copy
import xlsxwriter


def get_xlsx_file():
    files = os.listdir()
    ret = [file for file in files if file.endswith('.csv')]
    return ret

def create_output_file():
    xlsx = xlsxwriter.Workbook('清洗结果.xlsx')
    table = xlsx.add_worksheet('粗挑选')
    table.set_column('A:B', 50)
    table.write_string(0,0,'关键词')
    table.write_string(0,1,'搜索词')

    table2 = xlsx.add_worksheet('细挑选')
    table2.set_column('A:B', 50)
    table2.write_string(0,0,'关键词')
    table2.write_string(0,1,'搜索词')
    xlsx.close()

# 创建界面
def create_ui():
    #设置tkinter窗口
    root = Tk()
    root.geometry("500x400")
    root.resizable(width=False, height=False)
    root.title("搜索词完整搜索")

    #绘制两个label,grid（）确定行列
    Label(root,text="请输入计划名：").grid(row = 0,column =0)
    Label(root,text="请输入搜索词：").grid(row = 1,column =0)

    #导入两个输入框
    e1 = Entry(root)
    e2 = Entry(root)
    listBox = Text(root, width=40)

    #设置输入框的位置
    e1.grid(row =0 ,column =1)
    e2.grid(row =1 ,column =1)
    listBox.grid(row=3, column=1, columnspan=3)

    # 输入内容获取函数print打印
    def run():
        for filename in get_xlsx_file():
            if not e1.get():
                listBox.delete(1.0, END)
                listBox.insert(1.0, '清洗全部计划完成')
                select_words(filename, listBox)
            else:
                listBox.delete(1.0, END)
                listBox.insert(1.0,  '清洗{}计划完成'.format(e1.get()))
                select_words(filename, listBox, plan=e1.get())

    # 查询搜索词
    files = get_xlsx_file()
    data = pd.read_csv(files[0], encoding='gbk')
    data = data.drop_duplicates(subset=['搜索词'], keep='first')

    def search():
        if e2.get():
            content = data[data['搜索词'].str.contains(e2.get())][['关键词', '搜索词']]
            if len(content)<1:
                content = '没有对应搜索词'
            print(content)
            listBox.delete(1.0, END)
            listBox.insert(1.0, content)

    #设置两个按钮，点击按钮执行命令 command= 命令函数
    theButton1 = Button(root, text ="清洗计划报告", width =10,command =run)
    theButton2 = Button(root, text ="检索搜索词",width =10,command =search)

    #设置按钮的位置行列及大小
    theButton1.grid(row =0 ,column =3,sticky =W, padx=10,pady =5)
    theButton2.grid(row =1 ,column =3,sticky =E, padx=10,pady =5)

    mainloop()

# 读取过滤词
def read_filter_data():
    xl = xlrd.open_workbook('过滤词.xlsx')
    table = xl.sheets()[0]

    filter_long = []
    filter_short = []
    filter_char = []

    for label in range(3):
        if label==0:
            vector = filter_long
        elif label == 1:
            vector = filter_short
        else:
            vector = filter_char
        i = 2
        while True:
            try:
                word = table.row_values(i)[label]
                if type(word) == float:
                    word = str(int(word))
                if len(word)<1:
                    break
            except:
                break
            vector.append(word)
            i += 1

    return filter_long,filter_short,filter_char

# 获取数据集
def get_data(filename, listBox, rought=False, plan=False):

    data = pd.read_csv(filename, encoding='gbk')
    data = data[~data['搜索词'].str.contains('来自猜你喜欢')]
    if plan:
        data = data[data['推广计划'].str.contains(plan)]
        if len(data) == 0:
            listBox.delete(1.0, END)
            listBox.insert(1.0, '计划名找不到相关单元, 请检查计划名')
    # 1.筛选出含有手游和游戏的关键词
    if rought:
        data = data[data['搜索词'].str.contains('游戏')].append(data[data['搜索词'].str.contains('手游')]).append(data[data['匹配方式']=='目标客户追投'])
    # 2.搜索词去重
    data = data.drop_duplicates(subset=['搜索词'],keep='first')

    # 3.去除关键词游戏的多余名词
    filter_last = ['下载', '安卓', '手游', '游戏']
    for word in filter_last:
        data['关键词'] = data['关键词'].str.replace(word,'')

    return data

# 清理搜索词
def clean_words(data, filter_long, filter_short, filter_char):
    for word in filter_long:
        data['搜索词'] = data['搜索词'].str.replace(word, '')
    for word in filter_short:
        data['搜索词'] = data['搜索词'].str.replace(word, '')
    for word in filter_char:
        data['搜索词'] = data['搜索词'].str.replace(word, '')

    data = data.drop_duplicates(subset=['搜索词'], keep='first')
    return data

# 显示输出结果
def show_select_result(data, filename, rought=False):
    rought_select = []
    num = 0

    excel = xlrd.open_workbook('清洗结果.xlsx')
    new_excel = copy(excel)
    table = new_excel.get_sheet(0 if not rought else 1)

    for i in range(len(data)):
        if data.iloc[i][4] in data.iloc[i][5]:
            pass
        else:
            primary_word = data.iloc[i][4]
            search_word = data.iloc[i][5]
            if len(search_word)<3 and (search_word in primary_word):
                continue
            num += 1
            table.write(num, 0, primary_word)
            table.write(num, 1, search_word)
            line = '{}{}{}'.format(primary_word, ' '*(12-len(data.iloc[i][4])), search_word)
            print(line)
            # ???要放到continue前面吗
            rought_select.append(search_word)
    print('共{}个词'.format(num))
    new_excel.save('清洗结果.xlsx')

    return rought_select

# 粗细挑选
def select_words(filename, listBox, plan=False):
    create_output_file()
    filter_long, filter_short, filter_char = read_filter_data()
    # 4.粗挑选
    print('粗挑选结果:', '#'*30)
    data = get_data(filename, listBox, rought=True, plan=plan)
    clean_data = clean_words(data, filter_long, filter_short, filter_char)
    rought_select = show_select_result(clean_data, filename)

    # 6.细挑选
    print('细挑选结果', '#'*30)
    data = get_data(filename, listBox, plan=plan)
    data = clean_words(data, filter_long, filter_short, filter_char)

    # 6.1 去除包含粗调中的词
    for i in rought_select:
        if len(i)>2:
            data = data[~data['搜索词'].str.contains(i)]

    show_select_result(data, filename, rought=True)

if __name__ == '__main__':
    create_ui()
