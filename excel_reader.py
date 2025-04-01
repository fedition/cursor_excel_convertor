import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QLabel, QTextEdit, QFileDialog, QMessageBox,
                           QProgressBar, QTabWidget)
from PyQt5.QtCore import Qt, pyqtSignal
import pandas as pd
from data_processor import DataProcessor

class DropArea(QLabel):
    fileDropped = pyqtSignal(str)

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
            self.fileDropped.emit(files[0])
        event.accept()

class ExcelReader(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_processor = DataProcessor()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel文件读取器')
        self.setGeometry(100, 100, 800, 600)

        # 创建中央部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        # 创建拖放区域
        self.drop_area = DropArea()
        self.drop_area.setMinimumHeight(150)
        self.drop_area.fileDropped.connect(self.process_file)
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

        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 创建结果显示区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        layout.addWidget(self.result_text)

        # 创建数据预览标签页
        self.tab_widget = QTabWidget()
        self.tab_widget.setVisible(False)
        layout.addWidget(self.tab_widget)

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
        """处理Excel文件"""
        try:
            # 读取Excel文件
            excel_file = pd.ExcelFile(file_path)
            
            # 创建数据处理器实例
            processor = DataProcessor()
            processor.set_input_file_path(file_path)  # 设置输入文件路径
            
            # 处理每个工作表
            results = processor.process_all_data(excel_file)
            
            # 显示处理结果
            self.show_results(results)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理文件时出错：{str(e)}")

    def show_results(self, results):
        # 更新进度条
        self.progress_bar.setValue(100)
        
        # 显示结果
        result_text = "数据处理结果：\n\n"
        for sheet, result in results.items():
            result_text += f"{sheet}表：{result}\n"
        
        self.result_text.setText(result_text)
        
        # 显示处理后的数据
        self.show_processed_data()

    def show_processed_data(self):
        """显示处理后的数据"""
        self.tab_widget.clear()
        self.tab_widget.setVisible(True)

        # 显示商品数据
        if self.data_processor.products_df is not None:
            products_text = QTextEdit()
            products_text.setReadOnly(True)
            products_text.setText(str(self.data_processor.products_df))
            self.tab_widget.addTab(products_text, "商品数据")

        # 显示供应商数据
        if self.data_processor.suppliers_df is not None:
            suppliers_text = QTextEdit()
            suppliers_text.setReadOnly(True)
            suppliers_text.setText(str(self.data_processor.suppliers_df))
            self.tab_widget.addTab(suppliers_text, "供应商数据")

        # 显示库存数据
        if self.data_processor.inventory_df is not None:
            inventory_text = QTextEdit()
            inventory_text.setReadOnly(True)
            inventory_text.setText(str(self.data_processor.inventory_df))
            self.tab_widget.addTab(inventory_text, "库存数据")

        # 显示会员数据
        if self.data_processor.members_df is not None:
            members_text = QTextEdit()
            members_text.setReadOnly(True)
            members_text.setText(str(self.data_processor.members_df))
            self.tab_widget.addTab(members_text, "会员数据")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelReader()
    ex.show()
    sys.exit(app.exec_()) 