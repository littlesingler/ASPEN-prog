import xlwings as xw
import xlrd
import xlwt
import re
import numpy as np


# 修改xls格式文件单个单元格的值（不实用）
def change_cell_value(filename, sheet, box, cell_value):
    app = xw.App(visible=True,add_book=False)
    # 新建工作簿 (如果不接下一条代码的话，Excel只会一闪而过，卖个萌就走了）
    # wb = app.books.add()
    wb = app.books.open(filename)
    # 练习的时候建议直接用下面这条
    # wb = xw.Book('example.xlsx')
    # 这样的话就不会频繁打开新的Excel

    sht = wb.sheets[sheet]
    # sht = wb.sheets[第一个sheet名]
    old_value = sht.range(box).value
    print(old_value)
    sht.range(box).value = cell_value

    wb.save(filename)
    app.quit()


# 使用列表，修改xls格式文件多个单元格的值(多次操作的话，对于app打开和关闭特别麻烦，也影响效率，待改善)
def change_range_value(filename,sheet,info_list,change_flag=0):
    # sheet表示要修改的序号，info_list表示存有信息的列表，change_flag表示修改的标志：0，默认，表示修改
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(filename)
    sht = wb.sheets[sheet]
    box_list = ["c11","c14","c29","c33","f36","f37","f38"]    # 参数可修改
    for il in range(1,len(box_list)):
        if change_flag == 0:
            if info_list[1][il] != 0:      # 如果数据为空，证明不用修改
                sht.range(box_list[il-1]).value = info_list[1][il]

    wb.save()
    app.quit()


# 获取ASPEN里表名的序号
def get_sheet_seq(sheetname):
    # ASPEN里的表与其序号对应关系
    site_sheet = {"2D mark": 0, "FT1": 1, "SLT1": 2, "2nd Mark": 3, "vm": 4, "bake TR": 5,  "Pack-ppk": 6,
                  "2d mark": 0, "FT": 1,   "SLT": 2, "2nd MARK": 3, "Vm": 4, "bake TR ": 5, "pack": 6,
                  "2D MARK": 0,                      "2nd mark": 3, "VM": 4, "bake": 5,     "PACK": 6,
                  "2D": 0,                                                   "bake-TR ": 5, "Pack ppk": 6}
    return site_sheet[sheetname]
# s = get_sheet_seq("VM")
# print(s)


# 获取ASPEN表对应信息的位置(未完成)
def get_box(infor_name):
    # ASPEN的信息与表里的位置对应关系
    information_box_sheet0 = {"Test Code": 1, "Test Pgm": 2, "Accept Bins": 3, "NOT REQUIRED BIN": 4, "REJECT BIN": 5}
    information_box_sheet1 = {"PN":"c11","OPN":"c14","Test Code": "c29", "Test Pgm": "c33", "Accept Bins": "f36", "NOT REQUIRED BIN": "f37", "REJECT BIN": "f38"}
    information_box_sheet2 = {"PN":"c13","OPN":"c16","Test Code": "c29", "Test Pgm": "c33", "Accept Bins": "f36", "NOT REQUIRED BIN": "f37", "REJECT BIN": "f38"}
    information_box_sheet3 = {"Test Code": 1, "Test Pgm": 2, "Accept Bins": 3, "NOT REQUIRED BIN": 4, "REJECT BIN": 5}
    information_box_sheet4 = {"Test Code": 1, "Test Pgm": 2, "Accept Bins": 3, "NOT REQUIRED BIN": 4, "REJECT BIN": 5}
    information_box_sheet5 = {"Test Code": 1, "Test Pgm": 2, "Accept Bins": 3, "NOT REQUIRED BIN": 4, "REJECT BIN": 5}
    information_box_sheet6 = {"Test Code": 1, "Test Pgm": 2, "Accept Bins": 3, "NOT REQUIRED BIN": 4, "REJECT BIN": 5}
    if infor_name == "PN" or infor_name == "OPN" or infor_name == "Test Code" or infor_name == "Test Pgm" or infor_name == "Accept Bins" or infor_name == "NOT REQUIRED BIN" or  infor_name == "REJECT BIN":
        return information_box_sheet1[infor_name]
    else:
        print("输入有误")
# box = get_box("PN")
# print(box)  #注意当输入不正确时会返回none空值


# 从信息列表中获取所需要的信息(未完成)
def get_information(mode_info):
    # 数据模板
    sta_info_format = []
    site = ["2D making at TMP site", "MARK", "MARK VERIFICATION", "FT1","SLT1",
            "2ND MARK VERIFICATION","EXTERNAL VISUAL INSPECTION", "BAKE", "TNR PACK", "PACK",
            "OQC", "BOXSTOCK"]
    basic_info_format = ["T_site", "PN", "OPN", "Test Code", "Test Pgm",
                         "Accept Bins", "NOT REQUIRED BIN", "REJECT BIN"]
    for i in mode_info:
        print(i)
# get_information()


# 从excel模板读取信息,返回一个二维列表
def read_excel_infor(excel_file):
    # 0.打开excel操作
    data = xlrd.open_workbook(excel_file)
    # 1. 获取excel sheet对象
    table1 = data.sheets()[0]   # 模板放置在第一张表里
    rows = table1.nrows
    # col = table1.ncols
    # print("行数为%s \n列数为%s" % (rows, col))
    # 3. 获取整行和整列的数据.
    hangshuju = []
    for r in range(2, rows):
        row = table1.row_values(r)
        hangshuju.append(row)
    # col = table1.col_values(0)
    # print(hangshuju)
    # print(col)
    return hangshuju
# Excel_filename = r"C:\Users\xyue\Desktop\text.xlsx"  #要加r，转义字符会报错
# sj = read_excel_infor(Excel_filename)
# print(sj）


def wt_data_into_Aspen():
    pass


# 获取excel中某张表的数据总行数,sheet_number从0开始,返回数字类型
def get_row_count(excel_file,sheet_number):
    # 0.打开excel操作
    data = xlrd.open_workbook(excel_file)
    # 1. 获取excel sheet对象
    table1 = data.sheets()[sheet_number]  # 模板放置在第一张表里
    rows = table1.nrows
    return rows


# 查重,二维列表查重。判断一个二维列表中的一维列表，在已有表中的情况;返回不在列表里的数据list
def find_duplication(first_list,existed_list):
    # main_key = 0
    new_data = []
    for flx in first_list:
        if flx not in existed_list:
            new_data.append(flx)
            # print(str(flx)+"在列表里")
        # else:
            #new_data.append(flx)
    return new_data
# first_list = [[1,2,3],[1,3,2],[1,4,5],[3,9,9]]
# e_list = [[1,1,1],[1,0,0],[1,3,2],[1,4,5],[4,3,5],[1,2,3]]
# nl = find_duplication(first_list,e_list)
# print(nl)


# 去重，二维列表自身去重，返回一个二维列表
def duplication_checking(two_dimensional_list):
    arr = np.array(two_dimensional_list)
    return np.array(list(set([tuple(t) for t in arr])))


# 把记录添加进新表中(没有查重）
def remove_infor_for_store(readed_info,current_excel):
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(current_excel)
    sht = wb.sheets[1]
    row_count = get_row_count(current_excel,1)
    # zm = re.compile(r'[a-zA-Z]+').findall(box)[0]
    # sz = re.compile(r'\d+').findall(box)[0]
    # 用str转化总行数为字符串，好和range函数匹配
    # print("A"+str(row_count+1))

    # 在指定range添加一个二维列表信息
    sht.range("A"+str(row_count+1)).options(expand='table').value = readed_info
    wb.save(current_excel)
    app.quit()
    # print("运行完成")
# remove_infor_for_store(sj,Excel_filename)


aspen_filename = "2A5-0917024-00 NAV10GLXL.xls"
Excel_filename = r"C:\Users\xyue\Desktop\text.xlsx"

'''
if __name__ == '__main__':

    sj = read_excel_infor(Excel_filename)
    change_range_value(aspen_filename, 1, sj)
    print("更新数据完成")
    #filename = "2A5-0917024-00 NAV10GLXL.xls"

    #change_cell_value(filename,1,"c29","HELLO")'''
