# 创建Excel表格
# 1、安装openpyxl
# `pip install openpyxl`
# 官方文档 [](https://openpyxl.readthedocs.io/en/stable/)
# 官方中文文档 https://openpyxl-chinese-docs.readthedocs.io/zh_CN/latest/tutorial.html
# 参考文档 [最详细的Excel模块Openpyxl教程]https://zhuanlan.zhihu.com/p/342422919

from openpyxl.styles import Alignment
from openpyxl import Workbook
from openpyxl.styles import Font

from copy import copy

def createExcel(excelName, titles, posts):
    wb = Workbook()

    # grab the active worksheet
    ws = wb.active

    # 1、添加数据
    # for title in titles:
    # 1.1 添加第一行
    ws.append(titles)
    # 1.2 添加后续行
    for post in posts:
        ws.append([post.nickname, post.create_time, post.link])
    # 2、设置样式     
    # 设置单个列宽
    ws.column_dimensions['A'].width = 15.0
    ws.column_dimensions['B'].width = 30.0
    ws.column_dimensions['C'].width = 15.0
    ws.column_dimensions['D'].width = 15.0
    
    for col in ws.columns:
        for cell in col:
            # openpyxl styles aren't mutable,
            # so you have to create a copy of the style, modify the copy, then set it back
            if cell.column == 3:
                if cell.row > 1:
                    # 拿到指定的cell,设置样式
                    cell.hyperlink = cell.value # 设置为超链接
                    cell.font = Font(color='003366FF') # 设置字体颜色
            # 设置为居中
            alignment_obj = copy(cell.alignment)
            alignment_obj.horizontal = 'justify'
            alignment_obj.vertical = 'center'
            alignment_obj.wrap_text = True
            cell.alignment = alignment_obj

    # Save the file
    wb.save(excelName)

class Post:
    nickname = ""
    create_time = ""
    link = "http://www.baidu.com"

post1 = Post()
post1.nickname = "张三"
post1.create_time = "2022-10-29 19:00:00"

post2 = Post()
post2.nickname = "李四"
post2.create_time = "2022-10-29 19:00:00"

createExcel("table.xlsx", ["姓名","发布时间","链接"], [post1, post2])
