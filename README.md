# Word and Excel 操作

## 文件说明

- genTable.py
 + 利用python-docx，在word文档中生成一个年度目标的表格
 + 利用openpyxl，在excel中生成一个年度的目标表格

## 使用说明

- 安装 python-docx `pip install python-docx`
- 安装 openpyxl `pip install openpyxl`

## 修改说明

- 修改年份 `year = xxxx`
- 修改列名
 + Word 需要加一列
 + Excel直接在`row0`里面加列名


## 存在问题

- Word中单元格文字无法居中
- Excel 单元格无法自动调整大小 
