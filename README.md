# 提取表格（行）
用于提取多个表格文件（*.xlsx）中指定行的内容，并将这些内容汇集到新建的指定文件中。

## 依赖
1. python3
本工具基于 python3,需要首先安装

2. openpyxl
通过 pip 安装 openpyxl，如下： 
```sh 
$ pip install openpyxl 
```

## 使用方法
使用说明如下：
1. 将需要处理的 xlsx 文件全部集中到一个路径下，同时将本程序也置于相同路径。执行以下命令即可：
```sh
fetch_excel$ python3 fetch_rows.py  -o out.xlsx -s 3 -t 2 -p 1
```
其中，参数说明：\
-s  3：  标记提取行号为3，即：从第三行开始提取\
-t  2：  标记提取的标题栏位置为2 \
-p  1：  标记从文件的 sheet1 中提取

其余参数可以通过 -h 进行查看

```sh
fetch_excel$ python3 fetch_rows.py -h
 -s  指定起始行号，默认：1
 -e  指定结束列号，默认：1
 -o  指定输出文件名，默认：out.xlsx
 -p  指定 sheet，默认：1
 -t  指定标题行号,默认：1
```

