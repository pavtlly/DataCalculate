import xlwt


# 通用函数 获取列的最大值、最小值
def max_value(arr):
    val = float("-inf")
    for v in arr:
        if not isinstance(v, str):
            val = max(val, v)
    return val


def min_value(arr):
    val = float("inf")
    for v in arr:
        if not isinstance(v, str):
            val = min(val, v)
    return val


# 判断是否为负向指标
def check_negative(arr):
    for v in arr:
        if v == "负向指标":
            return True
    return False


# 获取红色字体
def get_red_font_style():
    font = xlwt.Font()
    font.bold = True  # 字体加粗
    font.colour_index = 2  # 红色
    style = xlwt.XFStyle()
    style.font = font
    return style


# 获取黄色背景
def get_yellow_background_style():
    font = xlwt.Font()
    font.colour_index = 0
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5
    style = xlwt.XFStyle()
    style.font = font
    style.pattern = pattern
    return style


# 获取红色背景
def get_red_background_style():
    font = xlwt.Font()
    font.colour_index = 0
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 2
    style = xlwt.XFStyle()
    style.font = font
    style.pattern = pattern
    return style


# 获取列的值总和
def get_col_sum(arr, col, start_row, end_row):
    sum_val = 0
    for i in range(start_row - 1, end_row):
        sum_val = sum_val + arr[i][col]
    return sum_val
