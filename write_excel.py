# coding=UTF-8
import xlrd
import xlwt
from xlutils.copy import copy


def write_excel_xls(path, sheet_name, value, style):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j], style=style)  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


def write_excel_xls_append(path, value, style):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    new_worksheet.write_merge(0, 0, 3, 5, "数学", style=style)
    new_worksheet.write_merge(0, 1, 0, 0, "题号", style=style)
    new_worksheet.write_merge(0, 1, 1, 1, "准考证号", style=style)
    new_worksheet.write_merge(0, 1, 2, 2, "姓名", style=style)
    new_worksheet.row(1).write(3, "得分", style)
    new_worksheet.row(1).write(4, "校次", style)
    new_worksheet.row(1).write(5, "班次", style)
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j], style=style)  # 追加写入数据，注意是从i+rows_old行开始写入
    # 设置单元格宽度
    new_worksheet.col(1).width = 256 * 20
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")


if __name__ == '__main__':
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    book_name_xls = r'C:\Users\j20687\Desktop\xls格式测试工作簿.xls'

    sheet_name_xls = 'xls格式测试表'

    value_title1 = [["", "", "", "", "职业", ""],
                    ["题号", "准考证号", "姓名", "得分", "校次", "班次"]]

    value_title2 = [
        ["", "", "", "", "职业", ""],
        ["题号", "准考证号", "姓名", "得分", "校次", "班次"],
        ["19", "123455667", "张三", "123", "232", "1"],
        ["22", "444444444", "李四", "111", "321", "12"],
        ["33", "343445455", "王五", "121", "212", "2"]
    ]

    value1 = [["19", "123455667", "张三", "123", "232", "1"],
              ["22", "444444444", "李四", "111", "321", "12"],
              ["33", "343445455", "王五", "121", "212", "2"]]

    write_excel_xls(book_name_xls, sheet_name_xls, value_title1, style)
    write_excel_xls_append(book_name_xls, value1, style)
