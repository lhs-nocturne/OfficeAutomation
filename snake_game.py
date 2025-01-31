import tkinter as tk
from tkinter import messagebox
import random

# 定义一些常量
CELL_SIZE = 20  # 每个格子的大小
WIDTH = 35  # 游戏区域的宽度（格子数）
HEIGHT = 40  # 游戏区域的高度（格子数）

# 设定通关分数
PASS_SCORE = 500

class SnakeGame:
    def __init__(self, master):  # 修改这里，接受一个主窗口作为参数
        # 使用 Toplevel 来创建二级窗口
        self.window = tk.Toplevel(master)
        self.window.title("贪吃蛇")
        self.window.resizable(False, False)         # 禁用窗口最大化
        self.canvas = tk.Canvas(self.window, width=WIDTH * CELL_SIZE, height=HEIGHT * CELL_SIZE)
        self.canvas.pack()

        # 让窗口获得焦点，解决需要鼠标点击才能控制的问题
        self.window.focus_force()

        # 将窗口放置在屏幕中间
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        window_width = WIDTH * CELL_SIZE
        window_height = HEIGHT * CELL_SIZE
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 初始化游戏相关变量
        self.score = 0
        self.snake = [(20, 20), (20, 21), (20, 22)]  # 蛇的初始位置
        self.direction = "Up"  # 蛇的初始方向
        self.food = self.generate_food()  # 生成食物
        self.game_over_flag = False
        # 初始速度，单位毫秒，数值越大速度越慢
        self.speed = 200

        # 绑定键盘事件
        self.window.bind("<Key>", self.on_key_press)

        # 开始游戏循环
        self.update()

    def generate_food(self):
        """生成食物的位置"""
        while True:
            x = random.randint(0, WIDTH - 1)
            y = random.randint(0, HEIGHT - 1)
            if (x, y) not in self.snake:
                return (x, y)

    def on_key_press(self, event):
        """处理键盘事件"""
        key = event.keysym
        if key == "Up" and self.direction != "Down":
            self.direction = "Up"
        elif key == "Down" and self.direction != "Up":
            self.direction = "Down"
        elif key == "Left" and self.direction != "Right":
            self.direction = "Left"
        elif key == "Right" and self.direction != "Left":
            self.direction = "Right"

    def move_snake(self):
        """移动蛇的位置"""
        head_x, head_y = self.snake[0]
        if self.direction == "Up":
            new_head = (head_x, head_y - 1)
        elif self.direction == "Down":
            new_head = (head_x, head_y + 1)
        elif self.direction == "Left":
            new_head = (head_x - 1, head_y)
        elif self.direction == "Right":
            new_head = (head_x + 1, head_y)

        # 判断是否撞到墙壁或自己
        if (
            new_head[0] < 0
            or new_head[0] >= WIDTH
            or new_head[1] < 0
            or new_head[1] >= HEIGHT
            or new_head in self.snake[1:]
        ):
            self.game_over_flag = True
        else:
            self.snake.insert(0, new_head)
            if new_head == self.food:
                self.score += 1
                # 每吃20分加快速度，最小速度为50毫秒
                if self.score % 20 == 0 and self.speed > 50:
                    self.speed -= 20
                self.food = self.generate_food()
            else:
                self.snake.pop()

    def draw_game(self):
        """绘制游戏界面"""
        self.canvas.delete("all")
        # 绘制蛇
        for segment in self.snake:
            x, y = segment
            self.canvas.create_rectangle(
                x * CELL_SIZE,
                y * CELL_SIZE,
                (x + 1) * CELL_SIZE,
                (y + 1) * CELL_SIZE,
                fill="green"
            )
        # 绘制食物
        x, y = self.food
        self.canvas.create_oval(
            x * CELL_SIZE,
            y * CELL_SIZE,
            (x + 1) * CELL_SIZE,
            (y + 1) * CELL_SIZE,
            fill="red"
        )
        # 绘制分数
        self.canvas.create_text(
            10, 10, text=f"分数: {self.score}", anchor="nw", font=("", 16)
        )

    def update(self):
        """游戏循环"""
        if not self.game_over_flag:
            self.move_snake()
            # 判断是否通关
            if self.score >= PASS_SCORE:
                self.pass_game()
            else:
                self.draw_game()
                self.window.after(self.speed, self.update)
        else:
            self.game_over()

    def game_over(self):
        """游戏结束处理"""
        messagebox.showinfo("提示", f"游戏结束，你的分数是: {self.score} 分")
        # 取消所有 after 调度任务
        self.window.after_cancel(self.update)
        self.window.destroy()

    def pass_game(self):
        """游戏通关处理"""
        messagebox.showinfo("恭喜通关", "你已经成功通关游戏！")
        # 取消所有 after 调度任务
        self.window.after_cancel(self.update)
        self.window.destroy()

def start_snake_game(master):  # 修改这里，接受一个主窗口作为参数
    game = SnakeGame(master)

# 以下代码是为了测试，可以在主程序中调用 start_snake_game 并传入主窗口实例
if __name__ == "__main__":
    root = tk.Tk()
    start_snake_game(root)
    root.mainloop()