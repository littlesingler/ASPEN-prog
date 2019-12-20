print("hello world!")
import pandas
import xlrd
import xlwt

# 0.打开excel操作
data = xlrd.open_workbook(r"C:\Users\xyue\Desktop\text.xlsx")



# 1. 获取excel sheet对象
table1 = data.sheets()[0]
table2 = data.sheet_by_index(0)
table3 = data.sheet_by_name(U"Sheet1")


rows=table1.nrows
col =table1.ncols
print("行数为%s \n列数为%s"%(rows,col))

#3. 获取整行和整列的数据.
row =table1.row_values(0)
col =table1.col_values(0)
print(row)
print(col)

# 4.获取单元格数据
cell_a1 = table1.cell_value(0, 0)
cell_x = table1.cell_value(1, 1)

print(cell_a1)
print(cell_x)


#  w1.创建workbook对象
#workbook = xlwt.Workbook(encoding ="utf-8",style_compression=0)

# 2.创建一个sheet对象,一个sheet对象对应excel文件中一张表格.
#sheet = workbook.sheet_index('FT')

#标题内容 = ["名字","年龄","学号"]
#记录1 = ["张三",18,10086]
#记录2 = ["李四",22,10010]

Aspen_file = "2A5-0917024-00 NAV10GLXL"



#def open_excel(file_name):
    #with open(file_name) as fn:
