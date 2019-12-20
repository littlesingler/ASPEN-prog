import pandas
import xlrd
import xlwt
from xlutils.copy import copy
import openpyxl
import traceback

filename = "2A5-0917024-00 NAV10GLXL.xls"

#获取Aspen文件单元格的值
def get_aspen_cell_value(filename,sheetname,box_fe,box_be,ret):
    oldWb = xlrd.open_workbook(filename)  #先打开已存在的表
    #oldWslist = oldWb.sheet_names()
    #print(oldWslist)
    oldWs = oldWb.sheet_by_name(sheetname)
    celldata = oldWs.cell_value(box_be-1,box_fe-1)

    #print(celldata)
    return celldata

#写入数据
def changeData(file, replaceText):
    # load the file(*.xlsx)
    wb = xlwt.Workbook
    # deal with one sheet
    ws = wb.worksheets[1]
    try:
        content = ws.cell(row=28, column=2).value
        print(content)
        if(content==replaceText):
            print("要替换的内容与原内容相同")
        else:
            ws.cell(row=28, column=2).value = replaceText
            wb.save(file)
    except Exception as e:
        print(traceback.format_exc())






#主函数
if __name__ == "__main__":

    aspen_cell_data = get_aspen_cell_value(filename,"FT1",3,29)
    #print(aspen_cell_data)



main()



#oldWs = oldWb.get_sheet(1)  #取sheet表
#print(oldWs)
#oldWs.write(28, 2, "pass")  #写入 2行4列写入pass
#oldWb.save() #保存至result路径
