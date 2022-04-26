import tkinter as tk
import xlrd
from tkinter.filedialog import *
from tkinter.messagebox import *
from calculate import *


# GUI相关函数
# 初始文件路径
def get_file_path():
    file_get = askopenfilename()  # 选择打开什么文件 返回文件名
    FilePath.set(file_get)


# 获取保存路径
def get_save_path():
    # 设置保存文件，并返回文件名，指定文件名后缀为.xls
    save_get = asksaveasfilename(defaultextension='.xls')
    SavePath.set(save_get)


# 读取Excel文件 sheet表
def load_data(file_path, excel_sheet):
    workbook = xlrd.open_workbook(str(file_path))  # excel路径
    sheet = workbook.sheet_by_name(excel_sheet)  # sheet表
    return sheet


# 执行计算模块
def exe_calculate_module():
    # noinspection PyBroadException
    try:
        sheet = load_data(str(FilePath.get()), str(read_sheet_name.get()))  # 读取sheet
        calculate(int(StartCol.get()), int(EndCol.get()),
                  int(StartRow.get()), int(EndRow.get()), sheet, str(SavePath.get()))
        showinfo(title='执行完毕', message=f'计算成功，结果保存在{str(SavePath.get())}')
    except IOError:
        showwarning(title='执行出错', message='无法执行计算，请检查所输入的信息')


# GUI界面部分
root = tk.Tk()
root.title('计算工具')
root.geometry('540x255')
FilePath = tk.StringVar()
SavePath = tk.StringVar()

# 选择文件、获取路径
tk.Label(root, text='选择文件：').place(x=20, y=20)
tk.Entry(root, textvariable=FilePath, width=50).place(x=90, y=20)
tk.Button(root, text='打开文件', command=get_file_path).place(x=460, y=15)

# 读取的sheet名称
tk.Label(root, text='读取Sheet：').place(x=20, y=70)
read_sheet_name = tk.Entry(root, width=50)
read_sheet_name.place(x=90, y=70)

# 起始、终止行列
tk.Label(root, text='起始列：').place(x=20, y=120)
StartCol = tk.Entry(root, width=5)
StartCol.place(x=70, y=120)
tk.Label(root, text='终止列：').place(x=130, y=120)
EndCol = tk.Entry(root, width=5)
EndCol.place(x=180, y=120)
tk.Label(root, text='起始行：').place(x=245, y=120)
StartRow = tk.Entry(root, width=5)
StartRow.place(x=295, y=120)
tk.Label(root, text='终止行：').place(x=355, y=120)
EndRow = tk.Entry(root, width=5)
EndRow.place(x=405, y=120)

# 保存路径
tk.Label(root, text='保存路径：').place(x=20, y=165)
tk.Entry(root, textvariable=SavePath, width=50).place(x=90, y=165)
tk.Button(root, text='选择目录', command=get_save_path).place(x=460, y=160)

# 运行按钮
tk.Button(root, text='开始计算', font=6, command=exe_calculate_module, width=10, height=1).place(x=180, y=200)

root.mainloop()