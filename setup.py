import sys
from cx_Freeze import setup, Executable

# 可执行文件的配置
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable(
        script="main.py",  # 主脚本文件名
        base=base,
        icon="static/app_icon.ico",  # 可选，指定图标文件路径
        target_name="办公自动化脚本整合.exe"# 指定打包后的可执行文件名称
    )
]

# 打包配置
setup(
    name="MyApp",  # 应用程序名称
    version="1.0",
    description="Your application description",  # 应用程序描述
    executables=executables,
    options={
        "build_exe": {
            "include_files": ["static"]
         }
    }

)