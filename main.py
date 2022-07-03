from openpyxl import load_workbook;
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import MergedCell


def parser_merged_cell(sheet: Worksheet, row, col):
    """
    Check whether it is a merged cell and get the value of the corresponding row and column cell.
    If it is a merged cell, take the value of the cell in the upper left corner of the merged area as the value of the current cell; otherwise, directly return the value of the cell
    : param sheet: current sheet object
    : param row: the row of the cell to be obtained
    : param col: the column of the cell to be obtained
    :return:
    """
    cell = sheet.cell(row=row, column=col)
    if isinstance(cell, MergedCell):  # judge whether the cell is a merged cell
        for merged_range in sheet.merged_cell_ranges:  # loop to find the merge range to which the cell belongs
            if cell.coordinate in merged_range:
                # Gets the cell in the upper left corner of the merge range and returns it as the value of the cell
                cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
    return cell


def get_xlsx(name, src_name, dst_name):
    wb = load_workbook(name)
    ws_data = wb.get_sheet_by_name(src_name)
    if dst_name in wb.sheetnames:
        ws_result = wb.get_sheet_by_name(dst_name)
    else:
        ws_result = wb.create_sheet(dst_name)
    row1 = ['测试案例编号', '计划测试日期', '前置条件', '测试概述', '操作步骤名称', '步骤描述', '预期结果', '状态']

    # 添加表头
    for index in range(len(row1)):
        _ = ws_result.cell(row=1, column=index + 1, value=row1[index])

    # 获取功能列表
    functions = []
    for row in ws_data.iter_rows(min_col=5, max_col=5):
        for cell in row:
            functions.append(cell.value)

    # 获取预期列表
    expected = []
    for row in ws_data.iter_rows(min_col=6, max_col=6):
        for cell in row:
            expected.append(cell.value)

    # 填写测试案例编号
    for rowIndex in range(len(functions)):
        if functions[rowIndex] is None:
            continue
        _ = ws_result.cell(row=rowIndex + 2, column=1, value='g00{}'.format(rowIndex))

    # 填写前置条件
    for rowIndex in range(len(functions)):
        if functions[rowIndex] is None:
            continue
        value = ''
        for column in ws_data.iter_cols(max_col=3, min_row=rowIndex + 1, max_row=rowIndex + 1):
            for cell in column:
                new_cell = parser_merged_cell(ws_data, cell.row, cell.column)
                if value == '':
                    value = new_cell.value
                    continue
                value = value + '->' + new_cell.value

        _ = ws_result.cell(row=rowIndex + 2, column=3, value=value)
        value = ''

    # 测试概述
    for rowIndex in range(len(functions)):
        if functions[rowIndex] is None:
            continue
        value = ''
        for column in ws_data.iter_cols(min_col=4, max_col=5, min_row=rowIndex + 1, max_row=rowIndex + 1):
            for cell in column:
                new_cell = parser_merged_cell(ws_data, cell.row, cell.column)
                if value == '':
                    value = '当' + new_cell.value + '时，'
                    continue
                value = value + '测试' + new_cell.value + '的情况'

        _ = ws_result.cell(row=rowIndex + 2, column=4, value=value)
        value = ''

    # 操作步骤名称
    for rowIndex in range(len(functions)):
        if functions[rowIndex] is None:
            continue
        _ = ws_result.cell(row=rowIndex + 2, column=5, value=functions[rowIndex])

    # 步骤描述
    for rowIndex in range(len(functions)):
        if functions[rowIndex] is None:
            continue
        value = ''
        for column in ws_data.iter_cols(min_col=4, max_col=5, min_row=rowIndex + 1, max_row=rowIndex + 1):
            for cell in column:
                new_cell = parser_merged_cell(ws_data, cell.row, cell.column)
                if value == '':
                    value = new_cell.value
                    continue
                value = value + '->' + new_cell.value

        _ = ws_result.cell(row=rowIndex + 2, column=6, value=value)
        value = ''

    # 预期结果
    for rowIndex in range(len(expected)):
        if expected[rowIndex] is None:
            continue
        _ = ws_result.cell(row=rowIndex + 2, column=7, value=expected[rowIndex])

    wb.save(name)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    fileList = [
        r"D:\ZSpaceFile\blockChain\userLogin.xlsx",
        r"D:\ZSpaceFile\blockChain\userOrder.xlsx",
    ]
    get_xlsx(fileList[1], 'userOrder', 'userOrderTest')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
