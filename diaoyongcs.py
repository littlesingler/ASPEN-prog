import ceshi
import change_xls_cell


# t = ceshi.asd
# print(t)
# ceshi.pri()

print(ceshi.__name__)
print()
print(__name__)
we = change_xls_cell.Excel_filename
print(we)

excel_name = r"C:\Users\xyue\Desktop\new bom.xlsx"
rws = change_xls_cell.get_row_count(excel_name,0)
print(rws)
print("完成")