import math
from func import *


# 输出初始值
def write_origin_info(start_col, end_col, start_row, end_row, sheet, worksheet):

    worksheet.write(end_row + 1, 0, u'标准化', get_red_font_style())  # 横坐标 纵坐标 内容 样式
    worksheet.write(2 * end_row + 1, 0, u'熵值', get_red_font_style())

    # 打印初始数据
    for i in range(0, end_row):
        for j in range(0, end_col):
            value = sheet.cell(i, j).value
            if value == "负向指标":
                worksheet.write(i, j, value, get_red_background_style())
            else:
                worksheet.write(i, j, value)

    # 打印列指标名称
    for i in range(start_row - 1, end_row):
        worksheet.write(i + end_row, 0, sheet.cell(i, 0).value)


# 起始列、终止列、起始行、终止行
def calculate(start_col, end_col, start_row, end_row, sheet, save_path):
    workbook = xlwt.Workbook(encoding='utf-8')  # 创建一个workbook 设置编码
    worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)  # 创建一个worksheet
    write_origin_info(start_col, end_col, start_row, end_row, sheet, worksheet)

    # 极差标准化运算
    standard_arr = [[0 for i in range(end_col + 10)] for i in range(end_row + 10)]
    extreme_sub = [0 for i in range(end_col + 10)]
    for j in range(start_col - 1, end_col):
        cols = sheet.col_values(j)
        cols_max = max_value(cols)  # 获取整列的最大值和最小值
        cols_min = min_value(cols)
        is_negative = check_negative(cols)
        extreme_sub[j] = cols_max - cols_min  # 计算极差
        for i in range(start_row - 1, end_row):
            pre_value = sheet.cell(i, j).value  # 获取对应行列单元格里数据
            if extreme_sub[j] == 0:  # 极差为0的情况
                # 正向指标且不全为0
                if (not is_negative) and (pre_value != 0):
                    value = 1
                else:
                    value = 0
                worksheet.write(i + end_row, j, value, get_yellow_background_style())
            else:
                if is_negative:
                    value = (cols_max - pre_value) / extreme_sub[j]  # 逆向指标值计算
                else:
                    value = (pre_value - cols_min) / extreme_sub[j]
                worksheet.write(i + end_row, j, value)
            standard_arr[i][j] = value
            standard_arr[end_row][j] = standard_arr[end_row][j] + value

    # 归一化运算
    normalize_ln_arr = [[0 for i in range(end_col + 10)] for i in range(end_row + 10)]
    for i in range(start_row - 1, end_row):
        for j in range(start_col - 1, end_col):
            if standard_arr[i][j] == 0:
                value = 0
                normalize_ln_arr[i][j] = value
            else:
                value = standard_arr[i][j] / standard_arr[end_row][j]
                normalize_ln_arr[i][j] = value * math.log(value)

    # 信息熵运算
    ex_info_entropy = [0 for i in range(end_col + 10)]  # 1-信息熵
    sum_value = 0
    for j in range(start_col - 1, end_col):
        value = -1 * get_col_sum(normalize_ln_arr, j, start_row, end_row) / math.log(end_row - start_row + 1)
        if value == 0 and extreme_sub[j] == 0 :
            ex_info_entropy[j] = 0
        else:
            ex_info_entropy[j] = 1 - value
        sum_value = sum_value + ex_info_entropy[j]

    for j in range(start_col - 1, end_col):
        value = ex_info_entropy[j] / sum_value
        worksheet.write(2 * end_row + 1, j, value)  # 写入对应单元格

    workbook.save(str(save_path))
