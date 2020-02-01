import openpyxl
from openpyxl import Workbook

import os
import sys
import getopt

# 逐一打印单元值 
# for row in sheet.rows:
#     for cell in row:
#         print( cell.value)
# 等于如下,
# 直接构造 list
# for row in rows:
#     line = [ col.value for col in row ]
def fetch_line( file_name ,page_index ,line_index):
    global start_line_index
    wb = openpyxl.load_workbook(cur_path+'/'+file_name)
    sheet_list = wb.get_sheet_names()

    print ( 'len,index:',len( sheet_list ) , (page_index-1))

    if len( sheet_list ) > (page_index-1) :
        sheet = wb.get_sheet_by_name( sheet_list[page_index-1] )
    else :
        print ( ' 异常：指定表索引超出范围' )
        return None

    if sheet.max_row >= line_index :
        rows = sheet.rows
        cnt = 0 
        for row in rows:
            cnt = cnt + 1
            if cnt == line_index:
                return [ col.value for col in row ]
    else :
        print('异常：指定行索引超出范围')
        return None

def append_line(sheet,line):
    if line != None:
        row = sheet.max_row + 1
        for i in  range(  len( line )  ):
            sheet.cell( row ,i+1).value = line[i]
    print (file,'......OK')


def usage( ):
    print(' -s  指定起始行号，默认：1')
    print(' -e  指定结束列号，默认：1')
    print(' -o  指定输出文件名，默认：out.xlsx')
    print(' -p  指定 sheet，默认：1')
    print(' -t  指定标题行号,默认：1')

## <<  程序入口 >> ##
opts,args = getopt.getopt( sys.argv[1:] ,'s:e:t:p:o:h')
start_line_index = 1
end_line_index = 1
fetch_sheet_index = 1
title_line_index = 1
out_file_name = 'out.xlsx'

for op,value in opts:
    if op == '-s':
        if value.isdigit():
            start_line_index = int(value)
        else :
            print('起始行号：无效字符')
    elif op == '-e':
        if value.isdigit():
            end_line_index = int(value)
        else :
            print('结束行号：无效字符')
    elif op == '-t':
        if value.isdigit():
            title_line_index = int(value)
        else :
            print('标题行：无效字符')
    elif op == '-p':
        if value.isdigit():
            fetch_sheet_index = int(value)
        else :
            print('Sheet：无效字符')
    elif op == '-o':
        out_file_name = value
    elif op == '-h':
        usage()
        sys.exit()

print( '标题行: %d\n指定表: %d\n起始行: %d \n结束行: %d\n输出文件: %s\n' % (title_line_index,fetch_sheet_index, start_line_index,end_line_index,out_file_name) )

## 查询是否有文件，如果有该文件，执行删除，再重新建同名文件  
cur_path  = sys.path[0]
file_name  = cur_path+'/'+out_file_name
if os.access( file_name , os.F_OK ):
    print ( 'File exist, remove it' )
    os.remove(file_name)

## 新建文件
w_wb = Workbook()
w_sheet_list = w_wb.get_sheet_names()
w_wb_sheet = w_wb.get_sheet_by_name(w_sheet_list[0])
w_wb_sheet = w_wb.active
w_wb.save(file_name)
print ( 'Create file %s , %s' % (file_name,w_sheet_list[0]) )

# 遍历全部文件
file_list = os.listdir(cur_path)
for file in file_list :
    # 遍历文件的范围：排除输出文件，仅检查xlsx文件，且不是 .~*.xlsx 临时文件
    if (( file != out_file_name )  and  ( file[-4:] == 'xlsx' ) and ( not file[:2] == '.~' )):
         # 从 file 中提取指定内容
         # 提取标题栏，，加入输出文件（仅执行一次）
        if title_line_index > 0 :
            line = fetch_line(  file ,fetch_sheet_index,title_line_index )
            append_line( w_wb_sheet , line )
            title_line_index = 0
        # 提取指定行，加入输出文件
        if end_line_index > start_line_index :
            # 多行提取
            for i in range( 0,end_line_index - start_line_index + 1):
                print('i=',int(i))
                line = fetch_line(  file , fetch_sheet_index , (start_line_index + int(i)) )
                append_line( w_wb_sheet , line )
        else :
            # 单行提取,没有设置 target_row1，则默认为1，小于等于 start_line_index
            line = fetch_line(  file , fetch_sheet_index , start_line_index )
            append_line( w_wb_sheet , line )

w_wb.save( file_name )