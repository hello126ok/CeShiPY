import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os

# 创建主窗口
root = tk.Tk()
root.title("Excel 行数据格式遍历处理器")
root.geometry("800x600")

# 全局变量，用于存储当前选择的文件路径
selected_file_path = ""

# 文件选择函数
def select_excel_file():
    global selected_file_path
    file_path = filedialog.askopenfilename(
        title="请选择一个 Excel 文件",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if file_path:
        selected_file_path = file_path
        file_label.config(text=f"已选择文件: {os.path.basename(file_path)}")
        process_excel_button.config(state=tk.NORMAL)  # 激活处理按钮
    else:
        messagebox.showwarning("警告", "未选择任何文件！")

# 处理 Excel 文件函数
def process_excel_file():
    global selected_file_path
    if not selected_file_path:
        messagebox.showerror("错误", "请先选择一个 Excel 文件！")
        return

    try:
        # 读取 Excel 文件中的所有 sheet
        excel_file = pd.ExcelFile(selected_file_path)
        sheet_names = excel_file.sheet_names
        output_text.delete(1.0, tk.END)  # 清空输出框

        output_text.insert(tk.END, f"开始处理 Excel 文件：{os.path.basename(selected_file_path)}\n")
        output_text.insert(tk.END, f"共包含 {len(sheet_names)} 个工作表（Sheet）:\n\n")

        all_sheets_data = {}

        for sheet_name in sheet_names:
            output_text.insert(tk.END, f"【工作表：{sheet_name}】\n")
            df = pd.read_excel(selected_file_path, sheet_name=sheet_name)

            # 获取行数与列数
            num_rows, num_cols = df.shape
            output_text.insert(tk.END, f"  - 行数：{num_rows}，列数：{num_cols}\n")
            output_text.insert(tk.END, f"  - 列名：{list(df.columns)}\n")

            output_text.insert(tk.END, "  - 开始逐行遍历数据：\n")

            # 遍历每一行数据（格式处理部分，这里只是打印，你可以自定义）
            for idx, row in df.iterrows():
                # 这里是“行数据格式的遍历处理”的地方，目前只是示例：打印行号与每行数据
                row_data = row.tolist()  # 转为 list
                row_display = f"    第 {idx + 1} 行: {row_data}\n"
                output_text.insert(tk.END, row_display)

                # 🔧 在此处添加你自己的“格式处理”逻辑，例如：
                # - 数据清洗
                # - 类型转换
                # - 条件判断
                # - 存储到新列表/字典/数据库等

            output_text.insert(tk.END, "\n")  # 工作表之间空一行

        output_text.insert(tk.END, "✅ Excel 文件所有工作表遍历完成！\n")

    except Exception as e:
        messagebox.showerror("处理错误", f"处理 Excel 文件时出错：{e}")
        output_text.insert(tk.END, f"❌ 处理出错：{e}\n")

# 创建界面组件
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

# 选择文件按钮
select_button = tk.Button(frame, text="选择 Excel 文件", command=select_excel_file, width=20)
select_button.pack(pady=5)

# 显示当前选中文件的标签
file_label = tk.Label(frame, text="未选择文件", fg="gray")
file_label.pack(pady=5)

# 处理按钮（一开始不可用）
process_excel_button = tk.Button(frame, text="开始处理 Excel 数据", command=process_excel_file, state=tk.DISABLED, width=20)
process_excel_button.pack(pady=5)

# 输出显示区域（带滚动条）
output_label = tk.Label(frame, text="处理结果/输出信息：")
output_label.pack(anchor=tk.W, pady=(10, 0))

output_text = scrolledtext.ScrolledText(frame, height=25, width=85, wrap=tk.WORD)
output_text.pack(fill=tk.BOTH, expand=True, pady=5)

# 启动主循环
root.mainloop()