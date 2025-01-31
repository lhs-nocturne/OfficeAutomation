import tkinter as tk
from split_excel_by_sheet import show_split_excel_by_sheet
from merge_excel_by_sheet import show_merge_excel_by_sheet
from merge_excel_by_first_sheet import show_merge_excel_by_first_sheet
from excel_to_csv import show_excel_to_csv
from PIL import Image, ImageTk
from relax import show_relax  # 导入贪吃蛇游戏启动函数

# =================使用windows高DPI支持 ================
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass
# ====================================================

# 记录当前选中的菜单按钮
current_selected = None

# 处理菜单按钮点击事件
def on_menu_click(button, func):
    global current_selected
    if current_selected:
        current_selected.config(bg="SystemButtonFace")
    button.config(bg="#C8C7C6")
    current_selected = button
    for widget in right_frame.winfo_children():
        widget.destroy()
    func()

# 首页对应的功能
def show_function_one():
    label = tk.Label(right_frame, text="您好，欢迎使用。", font=("", 14), fg="black")
    label.grid(row=0, column=0, pady=100, sticky=tk.NSEW)
    # label = tk.Label(right_frame, text="免责声明：使用本软件如若产生任何不良后果，软件作者不负任何法律责任。", font=("", 8))
    # label.grid(row=1, column=0, pady=200, sticky=tk.S)

# 创建主窗口
root = tk.Tk()
root.title("办公自动化脚本整合 v1.0")
# 设置窗口不可通过用户操作进行缩放
# 第一个参数为 False 表示禁止在水平方向上调整窗口大小
# 第二个参数为 False 表示禁止在垂直方向上调整窗口大小
root.resizable(False, False)

try:
    # 打开图片文件，替换为你的图片文件的实际路径
    icon_image = Image.open('static/app_icon.ico')
    photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, photo)
except Exception as e:
    print(f"未能加载图标文件: {e}")

# 获取屏幕的宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 设置窗口的宽度和高度
window_width = 1280
window_height = 860

# 计算窗口左上角的坐标，使窗口居中
x_coordinate = int((screen_width / 2) - (window_width / 2))
y_coordinate = int((screen_height / 2) - (window_height / 2))

# 设置窗口的位置和大小
root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

# 创建左侧菜单框架
left_frame = tk.Frame(root, bg="#E5E5E5")
left_frame.grid(row=0, column=0, sticky="ns")

# 创建右侧功能框架
right_frame = tk.Frame(root, relief='solid')
right_frame.grid(row=0, column=1, sticky="n", padx=50, pady=50)

# 设置网格布局的权重，使右侧框架可伸缩
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# 创建菜单按钮
menu_button1 = tk.Button(left_frame, width=16, bg="#E5E5E5", text="首页", command=lambda: on_menu_click(menu_button1, show_function_one))
menu_button1.grid(row=0, column=0, pady=0, sticky="ew")

menu_button2 = tk.Button(left_frame, width=16, bg="#E5E5E5", text="拆分Excel(页签版)", command=lambda: on_menu_click(menu_button2, lambda: show_split_excel_by_sheet(right_frame, root)))
menu_button2.grid(row=1, column=0, pady=0, sticky="ew")

menu_button3 = tk.Button(left_frame, width=16, bg="#E5E5E5", text="合并Excel(页签版)", command=lambda: on_menu_click(menu_button3, lambda: show_merge_excel_by_sheet(right_frame, root)))
menu_button3.grid(row=2, column=0, pady=0, sticky="ew")

menu_button4 = tk.Button(left_frame, width=16, bg="#E5E5E5", text="合并Excel(内容版)", command=lambda: on_menu_click(menu_button4, lambda: show_merge_excel_by_first_sheet(right_frame, root)))
menu_button4.grid(row=3, column=0, pady=0, sticky="ew")

menu_button5 = tk.Button(left_frame, width=16, bg="#E5E5E5", text="Excel批量转CSV", command=lambda: on_menu_click(menu_button5, lambda: show_excel_to_csv(right_frame, root)))
menu_button5.grid(row=4, column=0, pady=0, sticky="ew")

# 关于菜单按钮
menu_button_last = tk.Button(left_frame, width=16, bg="#E5E5E5", text="关于", command=lambda: on_menu_click(menu_button_last, lambda: show_relax(right_frame, root)))
menu_button_last.grid(row=5, column=0, pady=0, sticky="ew")

# 默认显示第一个功能，并模拟点击第一个菜单按钮的效果
# on_menu_click(menu_button1, show_function_one)
# 100毫秒后点击
root.after(100,lambda: on_menu_click(menu_button1, show_function_one))  # type: ignore

# 运行主事件循环
root.mainloop()