# excel2json
excel to json，将excel文件转换为json数组形式

## 依赖
openpyxl

    pip3 install openpyxl

## 配置
需要先配置config.py文件，确定输入输出文件名，以及sheet和header line
config.py  
    
    "input_file_path"       输入路径  
    "output_file_path"      输出路径  
    "sheet"                 要转换excel中的哪个sheet  
    "header_line"           header在哪一行，会从header行开始解析，跳过header之前的行  

## command:
python3 excel2json.py

## example input excel:  123.xlsx

|id|颜色|大小|性别|
|---|---|---|---|
|1|红|中|男|
|2|蓝||男|
|3||小|女|

## example output json:   out.json
``` json
[
  {
    "id":1,
    "颜色":"红",
    "大小":"中",
    "性别":"男",
  },
  {
    "id":2,
    "颜色":"蓝",
    "大小":null,
    "性别":"男",
  },
  {
    "id":3,
    "颜色":null,
    "大小":"小",
    "性别":"女",
  },
]
```
