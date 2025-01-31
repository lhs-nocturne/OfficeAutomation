import tkinter as tk
from snake_game import start_snake_game  # 导入贪吃蛇游戏启动函数
from game_2048 import start_2048_game  # 导入 2048 游戏启动函数
import webbrowser

def show_relax(frame, root):
    snake_button = tk.Button(frame, text="贪吃蛇（学习交流使用）", command=lambda: start_snake_game(root), width=30)
    snake_button.grid(row=0, column=0, padx=10, sticky=tk.EW, columnspan=3)
    # 添加开启 2048 游戏的按钮
    game_2048_button = tk.Button(frame, text="2048（学习交流使用）", command=lambda: start_2048_game(root), width=30)
    game_2048_button.grid(row=1, column=0, padx=10, sticky=tk.EW, columnspan=3)

    # 创建一个描述性文本的 Label
    description_label = tk.Label(frame, text="本软件完全开源，点击下面的链接访问源码：")
    description_label.grid(row=2, column=0, padx=10, pady=5)

    # 创建一个模拟超链接的 Label
    link_label = tk.Label(frame, text="https://github.com/lhs-nocturne/OfficeAutomation", fg="blue", cursor="hand2")
    link_label.grid(row=3, column=0, padx=10, pady=5)

    # 绑定鼠标点击事件
    link_label.bind("<Button-1>", lambda e: open_url("https://github.com/lhs-nocturne/OfficeAutomation"))

def open_url(url):
    webbrowser.open_new(url)