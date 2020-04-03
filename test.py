#coding=utf-8
import xlrd
import sys
sys.getdefaultencoding()!='utf-8'
# 打开Excel文件读取数据
data = xlrd.open_workbook('D:/workplace/金科研发4月3日.xlsx')
# 获取一个工作表
table = data.sheets()[0]                 #通过索引顺序获取
table = data.sheet_by_index(0)           #通过索引顺序获取
table = data.sheet_by_name(u'金科')       #通过名称获取
# 获取整行和整列的值（数组）
#print(table.col_values(5)[0])
#获取第一行内容 即标题
title = table.row_values(0)
print(title)
# 获取列数
num = table.nrows
for i in range(0,len(title)):
    if (title[i]=='早餐(5元)'):
        zaos=table.col_values(i);
        j=0
        list2=[]
        #index是列数
        for index in range(1,len(zaos)-1):
            if(zaos[index]=='是'):
                j=j+1
                list2.append(table.cell(index,0).value)
            s='早餐份数：'+repr(j)
            x='订早餐人:'+repr(list2)
        print(s)
        print(x)
    elif   (title[i]=='午餐(20元)'):
        zaos=table.col_values(i);
        j=0
        list2=[]
        #index是列数
        for index in range(1,len(zaos)-1):
            if(zaos[index]=='是'):
                j=j+1
                list2.append(table.cell(index,0).value)
            s='午餐份数：'+repr(j)
            x='订午餐人:'+repr(list2)
        print(s)
        print(x)
    elif   (title[i]=='晚餐(20元)'):
        zaos=table.col_values(i);
        j=0
        list2=[]
        #index是列数
        for index in range(1,len(zaos)-1):
            if(zaos[index]=='是'):
                j=j+1
                list2.append(table.cell(index,0).value)
            s='晚餐份数：'+repr(j)
            x='订晚餐人:'+repr(list2)
        print(s)
        print(x)
