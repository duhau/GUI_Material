# -*- coding: utf-8 -*-
#
# Created by: DuHua 2018-7-5
#
from tkinter import filedialog, messagebox
import os
import os.path
import sqlite3
import openpyxl

#针对每个xlsx文件的生成器
def eachXlsx(ZAID,Nuclide_name,Ion_density):
    for i in range(len(ZAID)):
        row=ZAID[i],Nuclide_name[i],Ion_density[i]
        yield tuple(map(lambda x:x.value, row))

#导入
def xlsx2sqlite(tabel_name,ZAID,Nuclide_name,Ion_density):
    #连接数据库，创建游标
    conn = sqlite3.connect('data.db')
    cur = conn.cursor()
    sql='CREATE TABLE IF NOT EXISTS %s(ZAID TEXT,NAME TEXT,ION_DENSITY REAL)'%tabel_name
    conn.execute(sql)
    #批量导入，减少提交事务的次数，可以提高速度
    sql1 = 'INSERT INTO %s VALUES(?,?,?)'%tabel_name
    cur.executemany(sql1, eachXlsx(ZAID,Nuclide_name,Ion_density))
    conn.commit()
    conn.close()
    

file_dir=os.getcwd()
filename = filedialog.askopenfilename(title='选择打开文件',initialdir=file_dir, filetypes = (("Excel文件", "*.xlsx") \
                                                                        ,("All files", "*.*") ))  
wb = openpyxl.load_workbook(filename,data_only=True)  #data_only=True可以将公式转化为数值
item=wb.get_sheet_names()  #获取所有的sheet名字
#print(item)
for file_name in item:
    ws = wb.get_sheet_by_name(file_name)
    # 遍历读取
    contents=[]
    for column in list(ws.columns):
        contents.append(column)
        
    ZAID=contents[0][3:]           #从第四行读取，A列
    Nuclide_name=contents[1][3:]   #从第四行读取，B列
    Ion_density=contents[13][3:]   #从第四行读取，N列
    xlsx2sqlite(file_name,ZAID,Nuclide_name,Ion_density)
    
    


