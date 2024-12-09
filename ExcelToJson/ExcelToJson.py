import os
import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import json

# 选择文件夹
def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()  # 打开文件夹选择框
    if folder_path:  # 如果文件夹路径不为空
        display_files(folder_path)  # 显示该路径下的xlsx文件

# 显示文件列表
def display_files(folder_path):
    # 获取该文件夹下的所有xlsx文件
    xlsx_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    # 清空之前的内容
    listbox.delete(0, tk.END)
    
    # 显示xlsx文件名
    if xlsx_files:
        for file in xlsx_files:
            listbox.insert(tk.END, file)
    else:
        listbox.insert(tk.END, "该文件夹没有.xlsx文件")

# 打开并读取选中的xlsx文件
def open_xlsx():
    selected_file = listbox.get(tk.ACTIVE)  # 获取当前选中的文件
    if selected_file:
        # 获取完整路径
        file_path = os.path.join(folder_path, selected_file)
        
        # 打开xlsx文件
        try:
            # 启动 Excel 应用
            app = xw.App(visible=False)  # 不显示 Excel 窗口
            wb = app.books.open(file_path)  # 打开工作簿
            sheet = wb.sheets[0]  # 获取第一个工作表
            
            # 获取总行数和总列数
            total_rows = sheet.used_range.rows.count  # 获取实际行数
            total_cols = sheet.used_range.columns.count  # 获取实际列数
            
            # 获取第一行的列标题
            column_titles = sheet.range(f'1:1').value  # 获取第一行数据（直接用value，不加[0]）
            
            # 读取第三行开始的每一行数据，并将字段名称替换为第一行的列标题
            data_array = []
            for row_idx in range(3, total_rows + 1):  # 从第三行开始
                row_data = sheet.range(f'{row_idx}:{row_idx}').value  # 获取整行数据
                
                # 如果这一行没有数据（所有单元格为空），则跳过
                if not any(row_data):  # 如果这一行的任何列的值都为空
                    continue
                
                # 确保只读取实际列数
                row_data = row_data[:total_cols]
                
                # 创建一个字典，将列标题作为键，行数据作为值
                row_dict = {column_titles[i]: row_data[i] for i in range(len(row_data))}
                data_array.append(row_dict)
            
            # 显示读取的数据
            label_data.config(text=f"前5行数据将在这里显示")
            label_rows_cols.config(text=f"总行数: {total_rows} 行\n总列数: {total_cols} 列")
            
            # 清空 Text 组件内容
            text_output.delete(1.0, tk.END)
            
            # 使用json.dumps来确保字段名称使用双引号
            data_output = "\n\n".join([json.dumps(item, ensure_ascii=False) for item in data_array[:5]])  # 只显示前5个对象
            text_output.insert(tk.END, f"读取的对象（前5行）:\n{data_output}")
            
            wb.close()  # 关闭工作簿
            app.quit()  # 退出 Excel 应用
        
        except Exception as e:
            label_data.config(text=f"无法打开文件：{e}")
            label_rows_cols.config(text="无法获取行列数")

# 创建窗口
root = tk.Tk()
root.title("Excel 导出 JSON 工具")

# 选择文件夹按钮
select_button = tk.Button(root, text="选择文件夹", command=select_folder)
select_button.pack(pady=10)

# 文件列表框
listbox = tk.Listbox(root, width=50, height=15)
listbox.pack(pady=10)

# 打开xlsx文件按钮
open_button = tk.Button(root, text="打开选中文件", command=open_xlsx)
open_button.pack(pady=10)

# 显示读取的数据
label_data = tk.Label(root, text="前5行数据将在这里显示", justify=tk.LEFT)
label_data.pack(pady=10)

# 显示总行数和总列数
label_rows_cols = tk.Label(root, text="总行数和总列数将在这里显示", justify=tk.LEFT)
label_rows_cols.pack(pady=10)

# 创建一个滚动条
scrollbar = tk.Scrollbar(root)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# 创建可交互的Text组件用于输出读取的对象
text_output = tk.Text(root, width=80, height=20, wrap=tk.WORD, yscrollcommand=scrollbar.set)
text_output.pack(pady=10)

# 配置滚动条与Text组件关联
scrollbar.config(command=text_output.yview)

# 运行主事件循环
root.mainloop()
