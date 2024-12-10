# Ver1.0 基本json对象输出

# import os
# import tkinter as tk
# from tkinter import filedialog
# import xlwings as xw
# import json

# # 选择文件夹
# def select_folder():
#     global folder_path
#     folder_path = filedialog.askdirectory()  # 打开文件夹选择框
#     if folder_path:  # 如果文件夹路径不为空
#         display_files(folder_path)  # 显示该路径下的xlsx文件

# # 显示文件列表
# def display_files(folder_path):
#     # 获取该文件夹下的所有xlsx文件
#     xlsx_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
#     # 清空之前的内容
#     listbox.delete(0, tk.END)
    
#     # 显示xlsx文件名
#     if xlsx_files:
#         for file in xlsx_files:
#             listbox.insert(tk.END, file)
#     else:
#         listbox.insert(tk.END, "该文件夹没有.xlsx文件")

# # 打开并读取选中的xlsx文件
# def open_xlsx():
#     selected_file = listbox.get(tk.ACTIVE)  # 获取当前选中的文件
#     if selected_file:
#         # 获取完整路径
#         file_path = os.path.join(folder_path, selected_file)
        
#         # 打开xlsx文件
#         try:
#             # 启动 Excel 应用
#             app = xw.App(visible=False)  # 不显示 Excel 窗口
#             wb = app.books.open(file_path)  # 打开工作簿
#             sheet = wb.sheets[0]  # 获取第一个工作表
            
#             # 获取总行数和总列数
#             total_rows = sheet.used_range.rows.count  # 获取实际行数
#             total_cols = sheet.used_range.columns.count  # 获取实际列数
            
#             # 获取第一行的列标题
#             column_titles = sheet.range(f'1:1').value  # 获取第一行数据（直接用value，不加[0]）
            
#             # 读取第三行开始的每一行数据，并将字段名称替换为第一行的列标题
#             data_array = []
#             for row_idx in range(3, total_rows + 1):  # 从第三行开始
#                 row_data = sheet.range(f'{row_idx}:{row_idx}').value  # 获取整行数据
                
#                 # 如果这一行没有数据（所有单元格为空），则跳过
#                 if not any(row_data):  # 如果这一行的任何列的值都为空
#                     continue
                
#                 # 确保只读取实际列数
#                 row_data = row_data[:total_cols]
                
#                 # 创建一个字典，将列标题作为键，行数据作为值
#                 row_dict = {column_titles[i]: row_data[i] for i in range(len(row_data))}
#                 data_array.append(row_dict)
            
#             # 显示读取的数据
#             label_data.config(text=f"前5行数据将在这里显示")
#             label_rows_cols.config(text=f"总行数: {total_rows} 行\n总列数: {total_cols} 列")
            
#             # 清空 Text 组件内容
#             text_output.delete(1.0, tk.END)
            
#             # 使用json.dumps来确保字段名称使用双引号
#             data_output = ",\n\n".join([json.dumps(item, ensure_ascii=False) for item in data_array[:5]])  # 只显示前5个对象
#             text_output.insert(tk.END, f"读取的对象（前5行）:\n{data_output}")
            
#             wb.close()  # 关闭工作簿
#             app.quit()  # 退出 Excel 应用
        
#         except Exception as e:
#             label_data.config(text=f"无法打开文件：{e}")
#             label_rows_cols.config(text="无法获取行列数")

# # 创建窗口
# root = tk.Tk()
# root.title("Excel 导出 JSON 工具")

# # 选择文件夹按钮
# select_button = tk.Button(root, text="选择文件夹", command=select_folder)
# select_button.pack(pady=10)

# # 文件列表框
# listbox = tk.Listbox(root, width=50, height=15)
# listbox.pack(pady=10)

# # 打开xlsx文件按钮
# open_button = tk.Button(root, text="打开选中文件", command=open_xlsx)
# open_button.pack(pady=10)

# # 显示读取的数据
# label_data = tk.Label(root, text="前5行数据将在这里显示", justify=tk.LEFT)
# label_data.pack(pady=10)

# # 显示总行数和总列数
# label_rows_cols = tk.Label(root, text="总行数和总列数将在这里显示", justify=tk.LEFT)
# label_rows_cols.pack(pady=10)

# # 创建一个滚动条
# scrollbar = tk.Scrollbar(root)
# scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# # 创建可交互的Text组件用于输出读取的对象
# text_output = tk.Text(root, width=80, height=20, wrap=tk.WORD, yscrollcommand=scrollbar.set)
# text_output.pack(pady=10)

# # 配置滚动条与Text组件关联
# scrollbar.config(command=text_output.yview)

# # 运行主事件循环
# root.mainloop()


# Ver 2.0 完整的json输出，支持自定义格式
import os
import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import json

class ExcelToJsonTool:
    def __init__(self, root):
        self.root = root
        self.root : tk.Tk
        self.file_paths = []
        self.setup_ui()

    def setup_ui(self):
        self.root.geometry("800x600")
        self.root.title("天狼星导表(ExToJs)小工具")

        # 设置窗口图标
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, 'Icon.ico')
        self.root.iconbitmap(icon_path)

        # 设置背景图片
        bg_image_path = os.path.join(current_dir, 'background.png')  # 背景图片路径
        if os.path.exists(bg_image_path):  # 确保背景图片存在
            bg_image = tk.PhotoImage(file=bg_image_path)
            bg_image
            bg_label = tk.Label(self.root, image=bg_image)
            bg_label.image = bg_image
            bg_label.place(relwidth=1, relheight=1)
        

        # 选择文件按钮
        select_button = tk.Button(self.root, text="选择 Excel 文件", command=self.select_files)
        select_button.pack(pady=10)

        # 文件列表框
        self.listbox = tk.Listbox(self.root, width=50, height=15)
        self.listbox.pack(pady=10)

        # 执行处理按钮
        process_button = tk.Button(self.root, text="开始处理文件", command=self.process_files)
        process_button.pack(pady=10)

        # 显示输出的文本
        self.text_output = tk.Text(self.root, width=80, height=20)
        self.text_output.pack(pady=10)

    def select_files(self):
        self.file_paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx")], title="选择一个或多个 Excel 文件"
        )
        if self.file_paths:
            self.display_files(self.file_paths)

    def display_files(self, file_paths):
        self.listbox.delete(0, tk.END)
        for file_path in file_paths:
            self.listbox.insert(tk.END, os.path.basename(file_path))

    def process_files(self):
        if not self.file_paths:
            self.text_output.delete(1.0, tk.END)
            self.text_output.insert(tk.END, "请先选择一个或多个文件！")
            return

        try:
            with xw.App(visible=False) as app:
                results = []
                for file_path in self.file_paths:
                    wb = app.books.open(file_path)
                    output_dir = os.path.dirname(file_path)

                    for sheet in wb.sheets:
                        sheet : xw.Sheet
                        if sheet.name == "备注":  # 跳过工作表名为“备注”的表
                            continue

                        total_rows = sheet.used_range.rows.count
                        total_cols = sheet.used_range.columns.count

                        column_titles = sheet.range(f'2:2').value
                        column_types = sheet.range(f'3:3').value

                        column_mapping = {column_titles[i]: column_types[i] for i in range(total_cols)}

                        data_array = []
                        for row_idx in range(4, total_rows + 1):
                            row_data = sheet.range(f'{row_idx}:{row_idx}').value
                            if not any(row_data):
                                continue
                            row_data = row_data[:total_cols]
                            row_dict = {}

                            for i in range(len(row_data)):
                                field_name = column_titles[i]
                                field_type = column_mapping[field_name]
                                value = row_data[i]

                                if value is None:
                                    row_dict[field_name] = None
                                else:
                                    try:
                                        if field_type == 'Int':
                                            row_dict[field_name] = int(value)
                                        elif field_type == 'Float':
                                            row_dict[field_name] = float(value)
                                        elif field_type == 'Int[]':
                                            row_dict[field_name] = [int(x) for x in str(value).split(',') if x.strip().isdigit()]
                                        elif field_type == 'Float[]':
                                            row_dict[field_name] = [float(x) for x in str(value).split(',') if x.strip()]
                                        elif field_type == 'Str':
                                            row_dict[field_name] = str(value)
                                        elif field_type == 'Str[]':
                                            row_dict[field_name] = [str(x.strip()) for x in str(value).split(',')]
                                        else:
                                            row_dict[field_name] = str(value)
                                    except Exception:
                                        row_dict[field_name] = str(value)

                            data_array.append(row_dict)

                        sheet_output = json.dumps(data_array, ensure_ascii=False, indent=4)
                        output_file_path = os.path.join(output_dir, f"{sheet.name}.json")
                        with open(output_file_path, 'w', encoding='utf-8') as f:
                            f.write(sheet_output)

                        results.append(f"{sheet.name}.json : 已导出到 {output_dir}")

                    wb.close()
                    del wb

                self.text_output.delete(1.0, tk.END)
                self.text_output.insert(tk.END, "\n".join(results))

        except Exception as e:
            self.text_output.delete(1.0, tk.END)
            self.text_output.insert(tk.END, f"错误：{e}")


if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelToJsonTool(root)
    root.mainloop()
