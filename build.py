import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 定义图标文件路径（如果有的话）
# icon_path = os.path.join(current_dir, 'icon.ico')

# 定义打包参数
params = [
    'excel_reader.py',  # 主程序文件
    '--name=Excel数据转换工具',  # 生成的exe名称
    '--noconsole',  # 不显示控制台窗口
    '--onefile',  # 打包成单个文件
    '--clean',  # 清理临时文件
    '--add-data=README.md;.',  # 添加README文件
    '--hidden-import=pandas',  # 添加必要的隐藏导入
    '--hidden-import=numpy',
    '--hidden-import=openpyxl',
    '--hidden-import=PyQt5',
    '--hidden-import=PyQt5.QtCore',
    '--hidden-import=PyQt5.QtGui',
    '--hidden-import=PyQt5.QtWidgets',
]

# 如果有图标文件，添加图标参数
# if os.path.exists(icon_path):
#     params.append(f'--icon={icon_path}')

# 执行打包命令
PyInstaller.__main__.run(params) 