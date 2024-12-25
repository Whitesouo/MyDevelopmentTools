import tkinter as tk
from tkinter import messagebox, Toplevel,scrolledtext
import xlwings as xw

# 当前的工具窗口引用
current_tool_window = None

def update_textbox_state():
    if var.get() == 'Text':
        entry_match.config(state='normal')
    else:
        entry_match.config(state='disabled')

def apply_rich_text():
    matched_text = entry_match.get()
    rich_text = entry_rich.get()
    
    try:
        color_hex = rich_text.split('=', 1)[1].split('>')[0]
        small_square.config(bg=color_hex)
    except IndexError:
        messagebox.showerror("Error", "Invalid color format")
        return
    
    selected_text = text_input.get("1.0", tk.END).strip()
    
    if var.get() == 'Text' and matched_text in selected_text:
        text_output.delete("1.0", tk.END)
        text_output.insert(tk.END, selected_text.replace(
            matched_text, f"{rich_text}{matched_text}</color>"))
    else:
        messagebox.showerror("Error", "No matching text found or invalid option")

def process_text():
    input_text = text_process_input.get()
    processed_text = text_output.get("1.0", tk.END).strip()
    text_process_output.delete("1.0", tk.END)
    text_process_output.insert(tk.END, f"{input_text}\n{processed_text}")

def open_rich_text_tool():
    tool_window = Toplevel(root)
    tool_window.title("富文本替换工具")
    tool_window.geometry("600x600")

    global var, entry_match, entry_rich, small_square, text_input, text_output
    global text_process_input, text_process_output

    var = tk.StringVar(value='Number')
    tk.Radiobutton(tool_window, text="数字", variable=var, value='Number', command=update_textbox_state).pack(anchor='w', padx=10, pady=(10, 0))
    tk.Radiobutton(tool_window, text="文本", variable=var, value='Text', command=update_textbox_state).pack(anchor='w', padx=10, pady=(5, 0))

    entry_match = tk.Entry(tool_window, state='disabled')
    entry_match.pack(anchor='w', padx=10, pady=(5, 0))

    tk.Label(tool_window, text="富文本信息:").pack(anchor='w', padx=10, pady=(10, 0))
    entry_rich = tk.Entry(tool_window)
    entry_rich.pack(anchor='w', padx=10, pady=(5, 0))

    small_square = tk.Label(tool_window, text="", bg="white", width=20, height=1)
    small_square.pack(anchor='w', padx=10, pady=(5, 10))

    text_input = tk.Text(tool_window, height=10)
    text_input.pack(anchor='w', padx=10, pady=(5, 0))

    btn_apply = tk.Button(tool_window, text="应用富文本替换", command=apply_rich_text)
    btn_apply.pack(anchor='w', padx=10, pady=(5, 10))

    text_output = tk.Text(tool_window, height=10)
    text_output.pack(anchor='w', padx=10, pady=(5, 0))

    text_process_input = tk.Entry(tool_window)
    text_process_input.pack(anchor='w', padx=10, pady=(5, 0))
    
    text_process_output = tk.Text(tool_window, height=5)
    text_process_output.pack(anchor='w', padx=10, pady=(5, 0))

    btn_process = tk.Button(tool_window, text="处理文本", command=process_text)
    btn_process.pack(anchor='w', padx=10, pady=(5, 10))

def concatenate_cells():
    try:
        # 连接到当前Excel应用
        app = xw.apps.active
        workbook = app.books.active
        worksheet = workbook.sheets.active
        selected_range = app.selection

        # 获取选中单元格的值并转换为整数（如果是数字类型），然后拼接
        values = []
        for cell in selected_range:
            value = cell.value
            if value is None:
                value = ''
            elif isinstance(cell.value, (int, float)):
                values.append(str(int(value)))
            else:
                values.append(str(value))
        
        concatenated_values = ','.join(values)

        # 将结果显示在Text控件中
        result_text.config(state=tk.NORMAL)
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, concatenated_values)
        result_text.config(state=tk.DISABLED)

        # 将结果拷贝到剪贴板
        root.clipboard_clear()
        root.clipboard_append(concatenated_values)
        messagebox.showinfo("Success", "Concatenated values copied to clipboard!")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def open_concatenate_tool():
    concatenate_window = Toplevel(root)
    concatenate_window.title("单元格拼接工具")
    concatenate_window.geometry("400x300")

    global result_text

    tk.Label(concatenate_window, text="请选择要拼接的Excel单元格").pack(pady=10)

    result_text = tk.Text(concatenate_window, height=10, state=tk.DISABLED)
    result_text.pack(padx=10, pady=5)

    btn_concatenate = tk.Button(concatenate_window, text="拼接选中单元格", command=concatenate_cells)
    btn_concatenate.pack(pady=10)

def open_rich_text_tool_remove():
    def remove_rich_text():
        input_text = text_input.get("1.0", tk.END)
        cleaned_text = remove_rich_text_tags(input_text)
        text_output.delete("1.0", tk.END)
        text_output.insert(tk.END, cleaned_text)
        copy_to_clipboard(cleaned_text)
    
    def remove_rich_text_tags(text):
        import re
        # 使用正则表达式去除 <color=#...> 和 </color> 标签 以及转义符
        clean_text = re.sub(r'<color=[^>]+>', '', text)
        clean_text = re.sub(r'</color>', '', clean_text)
        # 去除转义符 "\" 后面跟着的单个字符
        clean_text = re.sub(r'\\.', '', clean_text)
        return clean_text

    def copy_to_clipboard(text):
        rich_text_window.clipboard_clear()
        rich_text_window.clipboard_append(text)
        rich_text_window.update() # 保证剪贴板内容刷新
    
    rich_text_window = tk.Toplevel(root)
    rich_text_window.title("富文本替换工具")
    rich_text_window.geometry("600x600")
    
    tk.Label(rich_text_window, text="输入带有富文本代码的内容：").pack(pady=10)
    
    text_input = scrolledtext.ScrolledText(rich_text_window, wrap=tk.WORD, width=70, height=10)
    text_input.pack(pady=10)
    
    tk.Button(rich_text_window, text="移除富文本代码", command=remove_rich_text).pack(pady=10)
    
    tk.Label(rich_text_window, text="输出清理后的内容：").pack(pady=10)
    
    text_output = scrolledtext.ScrolledText(rich_text_window, wrap=tk.WORD, width=70, height=10)
    text_output.pack(pady=10)



def open_excel_data_extraction_tool():
    global current_tool_window

    if current_tool_window is not None:
        current_tool_window.destroy()

    def copy_to_clipboard(text):
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()

    def extract_and_display_data():
        column_input = column_entry.get().strip().upper()
        if not column_input:
            messagebox.showwarning("警告", "请输入需要提取的列")
            return

        columns = [col.strip() for col in column_input.split(',')]
        
        try:
            # 连接到当前Excel应用
            app = xw.apps.active
            workbook = app.books.active
            worksheet = workbook.sheets.active
            selected_range = app.selection

            if selected_range is None or selected_range.count == 0:
                raise Exception("没有选定任何单元格")

            # 刷新公式
            workbook.app.calculate()

            data = selected_range.value

            # 把数据转换成列表
            if not isinstance(data, list):
                data = [[data]]
            elif not isinstance(data[0], list):
                data = [data]

            # 转换列字母到索引
            column_indices = [ord(col) - ord('A') for col in columns]

            selected_data = []
            for row in data:
                extracted_row = [row[i] for i in column_indices if i < len(row)]
                selected_data.append(extracted_row)

            selected_data_str = "\n".join(["\t".join(map(str, row)) for row in selected_data])

            output_text.delete("1.0", tk.END)
            output_text.insert(tk.END, selected_data_str)
            copy_to_clipboard(selected_data_str)

        except Exception as e:
            messagebox.showerror("错误", f"出现错误: {e}")

    current_tool_window = tk.Toplevel(root)
    current_tool_window.title("Excel 行数据提取工具")
    current_tool_window.geometry("600x500")

    tk.Label(current_tool_window, text="请输入需要提取的列（例如：A 或 A,C）").pack(pady=10)
    column_entry = tk.Entry(current_tool_window, width=50)
    column_entry.pack(pady=5)

    tk.Button(current_tool_window, text="提取选定行数据", command=extract_and_display_data).pack(pady=10)

    output_text = scrolledtext.ScrolledText(current_tool_window, wrap=tk.WORD, width=50, height=20)
    output_text.pack(padx=10, pady=10)


def open_remove_line_break_tool():
    global current_tool_window

    if current_tool_window is not None:
        current_tool_window.destroy()

    def process_line_breaks():
        try:
            # 连接到当前Excel应用
            app = xw.apps.active
            workbook = app.books.active
            worksheet = workbook.sheets.active
            selected_range = app.selection

            if selected_range is None or selected_range.count == 0:
                raise Exception("没有选定任何单元格")

            # 遍历所选单元格
            for cell in selected_range:
                if isinstance(cell.value, str) and '\n' in cell.value:
                    cell.value = cell.value.replace('\n', '\\n')
                    cell.api.WrapText = False

            messagebox.showinfo("完成", "换行处理完成，所有单元格内换行已被替换为\\n，并关闭了自动换行")

        except Exception as e:
            messagebox.showerror("错误", f"出现错误: {e}")

    current_tool_window = tk.Toplevel(root)
    current_tool_window.title("Excel 换行处理工具")
    current_tool_window.geometry("400x200")

    tk.Label(current_tool_window, text="该工具将替换选定单元格中的换行符为\\n").pack(pady=20)
    tk.Button(current_tool_window, text="开始处理", command=process_line_breaks).pack(pady=10)




def create_gui():
    global root

    root = tk.Tk()
    root.title("天狼星Excel数据处理工具集")
    root.geometry("600x400")

    tk.Label(root, text="天狼星数据处理工具", font=("Helvetica", 18)).pack(pady=20)
    
    btn_open_rich_text = tk.Button(root, text="打开富文本替换工具", font=("Helvetica", 16), command=open_rich_text_tool)
    btn_open_rich_text.pack(pady=10)

    btn_open_concatenate = tk.Button(root, text="打开单元格拼接工具", font=("Helvetica", 16), command=open_concatenate_tool)
    btn_open_concatenate.pack(pady=10)
    
    btn_open_rich_text_remove = tk.Button(root, text="打开富文本剔除工具", font=("Helvetica", 16), command=open_rich_text_tool_remove)
    btn_open_rich_text_remove.pack(pady=10)
    
    btn_open_extraction_tool = tk.Button(root, text="打开Excel行数据提取工具", font=("Helvetica", 16), command=open_excel_data_extraction_tool)
    btn_open_extraction_tool.pack(pady=10)

    btn_open_remove_line_break_tool = tk.Button(root, text="打开Excel换行处理工具", font=("Helvetica", 16), command=open_remove_line_break_tool)
    btn_open_remove_line_break_tool.pack(pady=10)

    root.mainloop()

create_gui()