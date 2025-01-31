import tkinter as tk
from snake_game import start_snake_game  # 导入贪吃蛇游戏启动函数
from game_2048 import start_2048_game  # 导入 2048 游戏启动函数

def show_relax(frame, root):
    snake_button = tk.Button(frame, text="贪吃蛇（仅供学习交流使用）", command=lambda: start_snake_game(root), width=30)
    snake_button.grid(row=0, column=0, padx=10, sticky=tk.EW, columnspan=3)

    # 添加开启 2048 游戏的按钮
    game_2048_button = tk.Button(frame, text="2048（仅供学习交流使用）", command=lambda: start_2048_game(root), width=30)
    game_2048_button.grid(row=1, column=0, padx=10, sticky=tk.EW, columnspan=3)