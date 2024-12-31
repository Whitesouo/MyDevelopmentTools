import xlwings as xw
import keyboard
import tkinter as tk
from tkinter import messagebox
import os

# # 方法1: 改变选中单元格的字体颜色和加粗，除非是 ">"
# def Blue_color_and_bold_except_greater_than():
#     try:
#         # 连接到当前Excel应用
#         app = xw.apps.active
#         workbook = app.books.active
#         worksheet = workbook.sheets.active
        
#         selected_range = app.selection   # 获取当前选中的区域
        
#         # 如果没有选中任何单元格，抛出错误
#         if selected_range is None or selected_range.count == 0:
#             raise Exception("没有选定任何单元格")
        
#         # 获取单元格的值
#         cell_value = selected_range.value
        
#         # 如果单元格是空的，跳出处理
#         if cell_value is None:
#             raise Exception("单元格为空，无法处理")

#         # 将数字转为字符串，确保可以按字符处理
#         if isinstance(cell_value, (int, float)):
#             selected_range.font.color = (0, 112, 192) 
#             selected_range.font.bold = True
#             return

#         # 获取单元格文本内容的长度
#         text_length = len(cell_value)
#         print(text_length)
        
#         # 遍历单元格内的每个字符，进行字体修改
#         for i in range(text_length):
#             if cell_value[i] != '>':  # 如果字符不是 ">"
#                 # 使用字符范围 [i+1:i+2] 来选取当前字符并设置样式
#                 selected_range.characters[i :i + 1].font.color = (0, 112, 192)  # 修改颜色
#                 selected_range.characters[i :i + 1].font.bold = True  # 设置加粗

#             print(f"Text color and bold applied in {selected_range.address} except for '>' characters.")
        
#     except Exception as e:
#         # 弹出错误信息框
#         messagebox.showerror("Error", f"An error occurred: {e}")
#         print({e})

# 方法1: 改变选中单元格的字体颜色和加粗，除非是 ">"
def Blue_color_and_bold_except_greater_than():
    try:
        # 连接到当前Excel应用
        app = xw.apps.active
        workbook = app.books.active
        worksheet = workbook.sheets.active
        
        selected_range = app.selection  # 获取当前选中的区域
        
        # 如果没有选中任何单元格，抛出错误
        if selected_range is None or selected_range.count == 0:
            raise Exception("没有选定任何单元格")
        
        # 遍历选中的区域
        for cell in selected_range:
            # 获取单元格的值
            cell_value = cell.value
            
            # 如果单元格是空的，跳过
            if cell_value is None:
                continue

            # 如果是数字类型，直接处理单元格
            if isinstance(cell_value, (int, float)):
                cell.font.color = (0, 112, 192)  # 设置字体颜色为蓝色
                cell.font.bold = True  # 设置加粗
                continue

            # 如果单元格是字符串类型，则按字符处理
            if isinstance(cell_value, str):
                text_length = len(cell_value)
                
                # 遍历单元格内的每个字符，进行字体修改
                for i in range(text_length):
                    if cell_value[i] != '>':  # 如果字符不是 ">"
                        # 使用字符范围 [i+1:i+2] 来选取当前字符并设置样式
                        cell.characters[i :i + 1].font.color = (0, 112, 192)  # 修改颜色
                        cell.characters[i :i + 1].font.bold = True  # 设置加粗

            print(f"Text color and bold applied in {cell.address} except for '>' characters.")
        
    except Exception as e:
        # 弹出错误信息框
        messagebox.showerror("Error", f"An error occurred: {e}")
        print(f"Error: {e}")


# 方法2: 设置选中区域字体颜色为 RGB(75, 75, 75)，加斜体，取消加粗
def Commentize():
    try:
        # 连接到当前Excel应用
        app = xw.apps.active
        workbook = app.books.active
        worksheet = workbook.sheets.active

        selected_range = app.selection   # 假设选中A1单元格，可以根据需求修改
        
        if selected_range is None or selected_range.count == 0:
            raise Exception("没有选定任何单元格")
        
        selected_range.font.color = (75, 75, 75)  # 设置文字颜色
        selected_range.font.bold = False         # 设置加粗
        selected_range.font.italic = True        # 设置斜体
        # 获取当前字体大小并减小两次
        current_size = selected_range.font.size
        if current_size is not None:
            selected_range.font.size = current_size - 2  # 减少字体大小 2 个点
        print("Font color changed, bold removed, and italic applied.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# 方法3: 改变选中单元格的字体颜色变黄和加粗，除非是 ">"
def Yellow_color_and_bold_except_greater_than():
    try:
        # 连接到当前Excel应用
        app = xw.apps.active
        workbook = app.books.active
        worksheet = workbook.sheets.active
        
        selected_range = app.selection   # 获取当前选中的区域
        
        # 如果没有选中任何单元格，抛出错误
        if selected_range is None or selected_range.count == 0:
            raise Exception("没有选定任何单元格")
        
        # 遍历选中的区域
        for cell in selected_range:
            # 获取单元格的值
            cell_value = cell.value
            
            # 如果单元格是空的，跳过
            if cell_value is None:
                continue

            # 如果是数字类型，直接处理单元格
            if isinstance(cell_value, (int, float)):
                cell.font.color = (166, 101, 0)  # 设置字体颜色为蓝色
                cell.font.bold = True  # 设置加粗
                continue

            # 如果单元格是字符串类型，则按字符处理
            if isinstance(cell_value, str):
                text_length = len(cell_value)
                
                # 遍历单元格内的每个字符，进行字体修改
                for i in range(text_length):
                    if cell_value[i] != '>':  # 如果字符不是 ">"
                        # 使用字符范围 [i+1:i+2] 来选取当前字符并设置样式
                        cell.characters[i :i + 1].font.color = (166, 101, 0)  # 修改颜色
                        cell.characters[i :i + 1].font.bold = True  # 设置加粗
        
    except Exception as e:
        # 弹出错误信息框
        messagebox.showerror("Error", f"An error occurred: {e}")
        print({e})


# 设置快捷键绑定
def bind_shortcut_keys():
    # 定义快捷键
    keyboard.add_hotkey('ctrl+shift+d', Blue_color_and_bold_except_greater_than)  # 快捷键 CTRL+SHIFT+D
    keyboard.add_hotkey('ctrl+shift+c', Commentize)                               # 快捷键 CTRL+SHIFT+C
    keyboard.add_hotkey('ctrl+shift+e', Yellow_color_and_bold_except_greater_than)


# 创建UI窗口
def create_ui_window():
    window = tk.Tk()
    window.geometry("800x600")  # 设置窗口初始大小
    window.title("天狼星策划文档编写小工具")
    
    # 设置窗口图标为相对路径
    current_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前脚本所在目录
    icon_path = os.path.join(current_dir, 'Icon.ico')  # 构建相对路径
    window.iconbitmap(icon_path)  # 使用相对路径设置图标
    
    # 创建方法按钮
    method1_btn = tk.Button(window, text="标记为属性 (Ctrl+Shift+D)", command=Blue_color_and_bold_except_greater_than)
    method1_btn.pack(pady=10)
    
    method2_btn = tk.Button(window, text="标记为注释 (Ctrl+Shift+C)", command=Commentize)
    method2_btn.pack(pady=10)
    
    method3_btn = tk.Button(window, text="标记为功能 (Ctrl+Shift+E)", command=Yellow_color_and_bold_except_greater_than)
    method3_btn.pack(pady=10)

    # 显示窗口
    window.mainloop()

if __name__ == '__main__':
    # 在Excel中启动脚本时会调用这个部分
    bind_shortcut_keys()
    create_ui_window()
