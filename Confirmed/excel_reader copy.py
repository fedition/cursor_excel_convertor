import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QLabel, QTextEdit, QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt, pyqtSignal
import pandas as pd

class DropArea(QLabel):
    fileDropped = pyqtSignal(str)  # 添加信号

    def __init__(self):
        super().__init__()
        self.setAlignment(Qt.AlignCenter)
        self.setText("将Excel文件拖放到这里\n或者点击选择文件按钮")
        self.setStyleSheet("""
            QLabel {
                background-color: white;
                border: 2px dashed #cccccc;
                border-radius: 5px;
                padding: 20px;
                font-size: 14px;
            }
        """)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
            self.setStyleSheet("""
                QLabel {
                    background-color: #e1f5fe;
                    border: 2px dashed #2196f3;
                    border-radius: 5px;
                    padding: 20px;
                    font-size: 14px;
                    color: #1976d2;
                }
            """)
            self.setText("松开鼠标即可读取文件")
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            QLabel {
                background-color: white;
                border: 2px dashed #cccccc;
                border-radius: 5px;
                padding: 20px;
                font-size: 14px;
            }
        """)
        self.setText("将Excel文件拖放到这里\n或者点击选择文件按钮")

    def dropEvent(self, event):
        self.setStyleSheet("""
            QLabel {
                background-color: white;
                border: 2px dashed #cccccc;
                border-radius: 5px;
                padding: 20px;
                font-size: 14px;
            }
        """)
        self.setText("将Excel文件拖放到这里\n或者点击选择文件按钮")
        
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.fileDropped.emit(files[0])  # 发送信号
        event.accept()

class ExcelReader(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel文件读取器')
        self.setGeometry(100, 100, 600, 400)

        # 创建中央部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        # 创建拖放区域
        self.drop_area = DropArea()
        self.drop_area.setMinimumHeight(150)
        self.drop_area.fileDropped.connect(self.process_file)  # 连接信号
        layout.addWidget(self.drop_area)

        # 创建选择文件按钮
        self.select_button = QPushButton('选择文件')
        self.select_button.clicked.connect(self.select_file)
        self.select_button.setStyleSheet("""
            QPushButton {
                padding: 8px;
                font-size: 14px;
                background-color: #2196f3;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
        """)
        layout.addWidget(self.select_button)

        # 创建结果显示区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        layout.addWidget(self.result_text)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel files (*.xlsx);;All files (*.*)"
        )
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        try:
            # 读取Excel文件
            excel_file = pd.ExcelFile(file_path)
            
            # 需要读取的表名
            required_sheets = ['商品', '会员', '供应商', '库存']
            
            # 检查并读取每个表
            result_text = "文件读取结果：\n\n"
            for sheet in required_sheets:
                if sheet in excel_file.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    result_text += f"成功读取 '{sheet}' 表，共 {len(df)} 行数据\n"
                else:
                    result_text += f"未找到 '{sheet}' 表\n"
            
            self.result_text.setText(result_text)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取文件时发生错误：\n{str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelReader()
    ex.show()
    sys.exit(app.exec_()) 