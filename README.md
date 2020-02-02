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
fetch_excel$ python3 fetch_excel.py  -o out.xlsx -s 3 -t 2 -p 1
```
其中，参数说明：\
-s  3：  标记提取行号为3，即：从第三行开始提取\
-t  2：  标记提取的标题栏位置为2 \
-p  1：  标记从文件的 sheet1 中提取

其余参数可以通过 -h 进行查看

```sh
Usage: python3 fetch_excel [-s start-index] [-e end-index] [-m row/col] [-o *.xlsx] [-t title-index] [-p page-index]
 options:
 -s  提取行或列的起始位置，e.g. 1,2,3,..., def: 1
 -e  提取行或列的结束位置，e.g. 1,2,3,..., def: 1
 -o  指定汇总输出文件，e.g. my-out.xlsx ，def: out.xlsx
 -p  指定输入文件中的表序号，文件中第一张表对应为1，e.g. 1,2,3,..., def: 1
 -t  指定提取的标题行/列的位置,标题仅提取一次，e.g. 1,2,3,..., def: 1
 -m  指定提取模式，分为 row 和 col 模式, def : row
 Fetch Excel v1.0.0  2020/2/2 ( leoyang20102013@163.com )
```

