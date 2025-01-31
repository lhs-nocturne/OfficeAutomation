import tkinter as tk

def show_overlay(root):
    # 创建遮罩层
    overlay = tk.Frame(root, bg="#808080")  # 使用灰色背景模拟半透明效果
    overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
    # 创建提示文字
    label = tk.Label(overlay, text="任务正在执行中，请稍等...", font=("",16), bg="#808080", fg="white")
    label.place(relx=0.5, rely=0.5, anchor="center")

    return overlay