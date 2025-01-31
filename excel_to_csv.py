import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import threading
from mask_layer import show_overlay

def conversion_excel_files(source_folder, target_folder, root):
    # 封装好的遮罩层
    overlay = show_overlay(root)
    try:
        count = 0
        for root_dir, dirs, files in os.walk(source_folder):
            for file in files:
                if file.endswith(('.xls', '.xlsx')):
                    file_path = os.path.join(root_dir, file)
                    # 读取 Excel 文件
                    excel_file = pd.ExcelFile(file_path)
                    # 获取所有表名
                    sheet_names = excel_file.sheet_names
                    for sheet_name in sheet_names:
                        # 获取当前工作表的数据
                        df = excel_file.parse(sheet_name)
                        # 生成 CSV 文件名，包含表名
                        base_name = os.path.splitext(file)[0]
                        csv_file_name = f"{base_name}.csv"
                        csv_file_path = os.path.join(target_folder, csv_file_name)
                        df.to_csv(csv_file_path, index=False)
                    count += 1
        root.after(0, lambda: messagebox.showinfo("完成", f"成功转换文件 {str(count)} 个！"))
    except Exception as e:
        # 显示错误提示框
        root.after(0, lambda: messagebox.showerror("错误", f"合并过程中出现错误：{str(e)}"))
    finally:
        # 显示成功提示框
        overlay.destroy()

def show_excel_to_csv(frame, root):

    def select_source_folder():
        folder = filedialog.askdirectory()
        if folder:
            source_entry.delete(0, tk.END)
            source_entry.insert(0, folder)

    def select_target_folder():
        file = filedialog.askdirectory()
        if file:
            target_entry.delete(0, tk.END)
            target_entry.insert(0, file)

    def start_conversion():
        source_folder = source_entry.get()
        target_folder = target_entry.get()
        if not source_folder or not target_folder:
            messagebox.showwarning("提示", "请选择来源目录和目标目录！")
            return
        threading.Thread(target=conversion_excel_files, args=(source_folder, target_folder, root)).start()

    # 源文件夹输入框
    source_label = tk.Label(frame, text="来源目录:")
    source_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
    source_entry = tk.Entry(frame, width=35)
    source_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
    source_button = tk.Button(frame, text="浏览", command=select_source_folder)
    source_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

    # 目标文件输入框
    target_label = tk.Label(frame, text="目标目录:")
    target_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
    target_entry = tk.Entry(frame, width=35)
    target_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
    target_button = tk.Button(frame, text="浏览", command=select_target_folder)
    target_button.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)

    # 合并按钮
    merge_button = tk.Button(frame, text="执行转换", command=start_conversion)
    merge_button.grid(row=2, column=1, padx=10, pady=10,sticky=tk.EW)

    # 功能概述
    label_text = ("注：该功能为批量的将Excel转换为csv文件(支持.xlsx和xls后缀的文件)，"
                  "选择好Excel文件来源目录和转换成csv文件的目标目录，执行转换即可。")
    output_label = tk.Label(frame, text=label_text, justify=tk.LEFT, wraplength=627)
    output_label.grid(row=3, column=0, padx=5, pady=20, columnspan=3)