# Merge PPTX

一个用于合并PPTX的小工具，可以自动将多个PPTX文件合并成一个文件，并且将所有文字改为黑色，方便打印。

### 用法

将merge_pptx.py放到pptx的目录下，执行 ``` python merge_pptx.py ``` 即可。合并顺序为目录下pptx文件名的字典序，如果有特殊要求可以将文件名列表传入main函数，程序将按照给定的文件顺序合并。

### 依赖

+ Python3
+ python-pptx (pip install python-pptx)