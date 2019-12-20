

# 操作access数据库


# 用xlsxwriter向excel插入图片，存在问题:不能在原有的excel里插入
'''import xlsxwriter
ex = r'C:\Users\xyue\Desktop\text1.xlsx'
pic = r'C:\Users\xyue\Desktop\1.png'
y = xlsxwriter.Workbook(ex)    # 新建一个Workbook
ss = y.sheetnames
#print(ss)
h = y.add_worksheet(name='pic')
h.insert_image('A1', pic, {'x_scale': 0.3,'y_scale': 0.3,'url':"https://www.bilibili.com"})
sn = y.sheetnames
#print(sn)
y.close()'''


# xlwings插入图片，存在问题：不能调整图片大小和位置
'''import xlwings as xw
def change_range_value():
    ex = r'C:\Users\xyue\Desktop\text1.xlsx'
    pic = r'C:\Users\xyue\Desktop\1.png'
    app = xw.App(visible=True, add_book=False)
    wb = xw.Book(ex)
    #wb = app.books.open(ex)
    print(wb.sheets)
    # try:
    #     wb.sheets.add('hhh')
    # except ValueError as ve:
    #     print(ve)
    sht = wb.sheets[1]
    sht.pictures.add(pic)
    wb.save()
    app.quit()
change_range_value()'''


# 多返回值
'''def more_ret(lis):
    # for i in range(1,len(lis)):
        # return i,lis[i]   # return只返回一次
    l = []
    m = []
    for t in range(0,len(lis)):
        l.append(t)
        m.append(lis[t])
    return l,m,t      # 返回的是一个元组


lis =  [44,'er',2,544,'ww']
a = more_ret(lis)
print(a[1])'''

# 列表能进行比较
'''def __init__(self,st,nm):
   self.st = st
   self.nm = nm
d1 = [1,2,3]
d2 = [1,2,3]
d3 = [1,1,1]
if(d1 == d3):
    print("d1=d3")
else:
    print("no")
print(len(d1))
def pri():
    string1 = "er4332"
    string2 = "222str"
    print(string1)
a = 1
if a == 2:
    asd = 3
    print(asd)
ts = "dddddd"'''


# 可以读取公共盘文件
'''with open(r"\\ssuzfas01\P_BOM\CAMSTAR CONFIG\Xianping\前期事宜\笔记\yyy.txt") as fl:
    for line in fl:
        print(line)'''


# 未保存的excel记录，计算总行数时不算在内
'''from change_xls_cell import get_row_count
t = get_row_count(r"C:\Users\xyue\Desktop\text1.xlsx",0)
print(t)'''

#实现UI界面显示
'''
import sys
from PyQt5.QtWidgets import QWidget,QApplication,QMainWindow,QFormLayout
import simpleGUI


class SimpleDialogForm(simpleGUI.Ui_Form):#从自动生成的界面类继承
    def __init__(self, parent = None):
        super(SimpleDialogForm, self).__init__()
        
    def your_funcs(self):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = QMainWindow()          # 创建一个主窗体（必须要有一个主窗体）
    content = SimpleDialogForm()  # 创建对话框
    content.setupUi(main)         # 将对话框依附于主窗体
    main.show()                   # 主窗体显示
sys.exit(app.exec_())
'''

# if __name__ == '__main__':
#     fm = "jies"
#     #Shuchu.pri()
#     print()
#     pri()
#     print("结束")
# zm = re.compile(r'[a-zA-Z]+').findall(string1)[0]
# sz = re.compile(r'\d+').findall(string1)[0]
# print()
# print()
# n = 1
# s = [1,2,"a"]
# print(len(s))
# print(chr(ord(s) + 1))
