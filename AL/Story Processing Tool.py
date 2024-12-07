import tkinter as tk
from tkinter import messagebox

# 定义分类匹配规则
category_keywords = {
    "主线": "主线",
    "EX": "EX",
    "SP": "SP",
    "日常": "日常"
}

def process_data():
    # 获取输入的剧情数据
    input_text = text_input.get("1.0", tk.END).strip()
    selected_category = category_var.get()

    # 分类字典，用于存储匹配的活动标题
    matched_titles = []

    if not input_text:
        messagebox.showerror("错误", "请输入剧情数据！")
        return

    if selected_category not in category_keywords:
        messagebox.showerror("错误", "请选择一个有效的分类！")
        return

    # 按行分割输入文本
    lines = input_text.split('\n')

    for line in lines:
        # 检查是否包含选定的分类
        if selected_category in line[-6:]:
            try:
                # 提取第一个 | 和 第二个 | 之间的内容
                parts = line.split('|')
                if len(parts) >= 3:
                    title = parts[1].strip()  # 提取第一个和第二个 | 之间的部分
                    matched_titles.append(title)
            except ValueError:
                continue  # 如果没有找到标题格式，跳过这行

    # 将匹配到的标题输出到可复制的Label
    if matched_titles:
        label_output.config(state=tk.NORMAL)  # 启用文本框的编辑
        label_output.delete(1.0, tk.END)  # 清空现有内容
        label_output.insert(tk.END, "\n".join(matched_titles))  # 插入匹配的标题
        label_output.config(state=tk.DISABLED)  # 禁用文本框的编辑
    else:
        label_output.config(state=tk.NORMAL)
        label_output.delete(1.0, tk.END)
        label_output.insert(tk.END, "没有匹配到任何标题。")
        label_output.config(state=tk.DISABLED)

# 创建窗口
root = tk.Tk()
root.title("剧情活动分类提取工具")

# 输入框：用于输入剧情信息数据
label_input = tk.Label(root, text="输入剧情数据：")
label_input.pack(pady=5)

text_input = tk.Text(root, height=15, width=60)
text_input.pack(pady=5)

# 分类选项：选择需要提取的活动类型
label_category = tk.Label(root, text="选择活动类型：")
label_category.pack(pady=5)

category_var = tk.StringVar()
category_menu = tk.OptionMenu(root, category_var, "主线", "EX", "SP", "日常")
category_menu.pack(pady=5)

# 按钮：处理数据并提取标题
process_button = tk.Button(root, text="提取标题", command=process_data)
process_button.pack(pady=10)

# 可复制的文本框：显示提取到的活动标题
label_output = tk.Text(root, height=10, width=60, wrap=tk.WORD, state=tk.DISABLED)
label_output.pack(pady=10)

# 启动窗口
root.mainloop()
