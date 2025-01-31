import tkinter as tk
import random  # 导入random模块

class Game2048:
    def __init__(self, master):
        # 使用 Toplevel 来创建二级窗口
        self.window = tk.Toplevel(master)
        self.window.title("2048")

        # 禁用窗口最大化
        self.window.resizable(False, False)

        # 将窗口放置在屏幕中间
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        # 增大尺寸
        window_width = 800
        window_height = 800
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建画布，尺寸也增大两倍
        self.canvas = tk.Canvas(self.window, width=800, height=800)
        self.canvas.pack()

        # 初始化游戏相关变量
        self.board = [[0] * 4 for _ in range(4)]
        self.add_random_tile()
        self.add_random_tile()

        # 绘制游戏界面
        self.draw_board()

        # 绑定键盘事件
        self.window.bind("<Key>", self.on_key_press)
        self.window.focus_set()  # 让窗口获得焦点

        # 新增标志位，用于记录是否已经提示过 2048 消息
        self.has_shown_2048_message = False

    def add_random_tile(self):
        """在空白格子中随机添加一个 2 或 4 的方块"""
        empty_cells = [(i, j) for i in range(4) for j in range(4) if self.board[i][j] == 0]
        if empty_cells:
            i, j = random.choice(empty_cells)
            # 90% 的概率生成 2，10% 的概率生成 4
            self.board[i][j] = 2 if random.random() < 0.9 else 4

    def draw_board(self):
        """绘制游戏界面"""
        self.canvas.delete("all")
        # 单元格尺寸增大两倍
        cell_size = 200
        for i in range(4):
            for j in range(4):
                value = self.board[i][j]
                x = j * cell_size
                y = i * cell_size
                if value != 0:
                    # 绘制数字方块，带有浅灰色边缘
                    self.canvas.create_rectangle(x, y, x + cell_size, y + cell_size, fill="lightgray")
                    self.canvas.create_rectangle(x + 5, y + 5, x + cell_size - 5, y + cell_size - 5, fill=self.get_color(value))
                    # 根据数字长度动态调整字体大小
                    font_size = self.get_font_size(value)
                    self.canvas.create_text(x + cell_size / 2, y + cell_size / 2, text=str(value), font=("", font_size))

    def get_color(self, value):
        """根据方块的值返回对应的颜色"""
        colors = {
            2: "#eee4da", 4: "#ede0c8", 8: "#f2b179", 16: "#f59563",
            32: "#f67c5f", 64: "#f65e3b", 128: "#edcf72", 256: "#edcc61",
            512: "#edc850", 1024: "#edc53f", 2048: "#edc22e"
        }
        return colors.get(value, "#cdc1b4")

    def get_font_size(self, value):
        """根据数字的长度返回合适的字体大小"""
        if value < 10:
            return 64
        elif value < 100:
            return 56
        elif value < 1000:
            return 48
        else:
            return 32

    def on_key_press(self, event):
        """处理键盘事件"""
        key = event.keysym
        if key in ["Up", "Down", "Left", "Right"]:
            original_board = [row[:] for row in self.board]  # 记录原始的棋盘状态
            moved = False
            if key == "Up":
                moved = self.move_up()
            elif key == "Down":
                moved = self.move_down()
            elif key == "Left":
                moved = self.move_left()
            elif key == "Right":
                moved = self.move_right()

            if moved:
                self.add_random_tile()
                self.draw_board()

                # 检查是否合成了 2048，且还未提示过
                if not self.has_shown_2048_message:
                    for i in range(4):
                        for j in range(4):
                            if self.board[i][j] == 2048:
                                tk.messagebox.showinfo("恭喜！", "恭喜！完成2048")
                                self.has_shown_2048_message = True
                                break

            # 如果没有移动，恢复原始棋盘状态
            elif self.board != original_board:
                self.board = original_board
                self.draw_board()

            # 检查是否无法移动，如果无法移动则提示游戏结束
            if not self.can_move():
                tk.messagebox.showinfo("游戏结束", "无法移动，游戏结束！")

    def move_up(self):
        """处理向上移动的逻辑"""
        moved = False
        for col in range(4):
            new_col = [self.board[row][col] for row in range(4) if self.board[row][col] != 0]
            merged_col = []
            skip = False
            for i in range(len(new_col)):
                if skip:
                    skip = False
                    continue
                if i < len(new_col) - 1 and new_col[i] == new_col[i + 1]:
                    merged_col.append(new_col[i] * 2)
                    skip = True
                    moved = True
                else:
                    merged_col.append(new_col[i])
            merged_col += [0] * (4 - len(merged_col))
            for row in range(4):
                if self.board[row][col] != merged_col[row]:
                    moved = True
                self.board[row][col] = merged_col[row]
        return moved

    def move_down(self):
        """处理向下移动的逻辑"""
        moved = False
        for col in range(4):
            new_col = [self.board[row][col] for row in range(3, -1, -1) if self.board[row][col] != 0]
            merged_col = []
            skip = False
            for i in range(len(new_col)):
                if skip:
                    skip = False
                    continue
                if i < len(new_col) - 1 and new_col[i] == new_col[i + 1]:
                    merged_col.append(new_col[i] * 2)
                    skip = True
                    moved = True
                else:
                    merged_col.append(new_col[i])
            merged_col += [0] * (4 - len(merged_col))
            for row in range(3, -1, -1):
                if self.board[row][col] != merged_col[3 - row]:
                    moved = True
                self.board[row][col] = merged_col[3 - row]
        return moved

    def move_left(self):
        """处理向左移动的逻辑"""
        moved = False
        for row in range(4):
            new_row = [self.board[row][col] for col in range(4) if self.board[row][col] != 0]
            merged_row = []
            skip = False
            for i in range(len(new_row)):
                if skip:
                    skip = False
                    continue
                if i < len(new_row) - 1 and new_row[i] == new_row[i + 1]:
                    merged_row.append(new_row[i] * 2)
                    skip = True
                    moved = True
                else:
                    merged_row.append(new_row[i])
            merged_row += [0] * (4 - len(merged_row))
            for col in range(4):
                if self.board[row][col] != merged_row[col]:
                    moved = True
                self.board[row][col] = merged_row[col]
        return moved

    def move_right(self):
        """处理向右移动的逻辑"""
        moved = False
        for row in range(4):
            new_row = [self.board[row][col] for col in range(3, -1, -1) if self.board[row][col] != 0]
            merged_row = []
            skip = False
            for i in range(len(new_row)):
                if skip:
                    skip = False
                    continue
                if i < len(new_row) - 1 and new_row[i] == new_row[i + 1]:
                    merged_row.append(new_row[i] * 2)
                    skip = True
                    moved = True
                else:
                    merged_row.append(new_row[i])
            merged_row += [0] * (4 - len(merged_row))
            for col in range(3, -1, -1):
                if self.board[row][col] != merged_row[3 - col]:
                    moved = True
                self.board[row][col] = merged_row[3 - col]
        return moved

    def can_move(self):
        """检查是否还能移动"""
        # 检查是否有空白格子
        for i in range(4):
            for j in range(4):
                if self.board[i][j] == 0:
                    return True

        # 检查相邻格子是否有相同的值
        for i in range(4):
            for j in range(3):
                if self.board[i][j] == self.board[i][j + 1]:
                    return True
        for i in range(3):
            for j in range(4):
                if self.board[i][j] == self.board[i + 1][j]:
                    return True

        return False

def start_2048_game(master):
    game = Game2048(master)


# 以下是测试代码，可以在主程序中调用 start_2048_game 来启动游戏
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    start_2048_game(root)
    root.mainloop()