import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog, QTableView, QWidget

class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 初始化窗口
        self.setGeometry(300, 300, 600, 400)
        self.setWindowTitle('Excel-SQL Viewer')

        # 创建标签页
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # 添加显示数据标签页
        self.data_tab = QWidget()
        self.tabs.addTab(self.data_tab, "View Data")

        # 设置显示数据标签页的布局
        self.data_tab_layout = QVBoxLayout()
        self.data_tab.setLayout(self.data_tab_layout)

        # 添加文件选择按钮
        self.btn_open = QPushButton("Open Excel File")
        self.btn_open.clicked.connect(self.openFileDialog)
        self.data_tab_layout.addWidget(self.btn_open)

        # 添加文件路径文本框
        self.line_edit = QLineEdit()
        self.data_tab_layout.addWidget(self.line_edit)

        # 添加表格视图
        self.table_view = QTableView()
        self.data_tab_layout.addWidget(self.table_view)

        self.show()

    def openFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
        if fileName:
            self.line_edit.setText(fileName)
            self.loadExcel(fileName)

    def loadExcel(self, file_path):
        self.df = pd.read_excel(file_path)
        # 将DataFrame显示在表格视图中...

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelViewer()
    sys.exit(app.exec_())
