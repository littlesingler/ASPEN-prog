
# import pypyodbc
# import win32com.client
# path=r'C:\Users\xyue\Desktop\TEXT.accdb'# 数据库文件
#
# con = win32com.client.Dispatch(r'ADODB.Connection')
# DSN = 'PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=' + path + ';'
# con.Open(DSN)
# rs = win32com.client.Dispatch(r'ADODB.Recordset')
# rs.Cursorlocation = 3
# rs.Open('SELECT TOP 1 * FROM table1', con)
# for i in range(0, rs.Fields.Count):
#     print(rs.Fields[i].Name + ' | ' + str(rs.Fields[i].Type) + ' | ' + str(rs.Fields[i].DefinedSize))

# -*-coding:utf-8-*-
# access数据库连接成功
import pyodbc

# 连接数据库（不需要配置数据源）,connect()函数创建并返回一个 Connection 对象
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\xyue\Desktop\TEXT.accdb')
# cursor()使用该连接创建（并返回）一个游标或类游标的对象
crsr = cnxn.cursor()

# 打印数据库goods.mdb中的所有表的表名
print('`````````````` goods ``````````````')
# t = crsr.tables(tableType='TABLE')
# print(t)
# >>> <pyodbc.Cursor object at 0x010690E0>
for table_info in crsr.tables(tableType='TABLE'):
    print(table_info.table_name)


# 提交数据（只有提交之后，所有的操作才会对实际的物理表格产生影响）
crsr.commit()
crsr.close()
cnxn.close()
