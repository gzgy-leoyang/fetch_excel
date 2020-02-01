# 提取表格（行）


## 依赖
### python3

### openpyxl

## 使用方法
使用说明如下：
```sh
fetch_excel$ python3 fetch_rows.py -h
 -s  指定起始行号，默认：1
 -e  指定结束列号，默认：1
 -o  指定输出文件名，默认：out.xlsx
 -p  指定 sheet，默认：1
 -t  指定标题行号,默认：1
```

执行以下命令即可
```sh
fetch_excel$ python3 fetch_rows.py -o out.xlsx -s 3 -t 2 -p 1
```