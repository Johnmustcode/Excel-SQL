import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog, QTableView, QWidget
from PyQt5.QtCore import QAbstractTableModel, Qt

class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data  # DataFrame数据

    def rowCount(self, parent=None):
        return self._data.shape[0]  # 返回行数

    def columnCount(self, parent=None):
        return self._data.shape[1]  # 返回列数

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])  # 返回单元格数据
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._data.columns[section])  # 返回列标题
        else:
            return str(self._data.index[section])  # 返回行标题
        
class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 初始化窗口
        self.setGeometry(300, 300, 600, 400)
        self.setWindowTitle('Excel-SQL Viewer')  # 设置窗口标题

        # 创建标签页
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # 添加显示数据标签页
        self.data_tab = QWidget()
        self.tabs.addTab(self.data_tab, "View Data")  # 创建数据查看页签

        # 设置显示数据标签页的布局
        self.data_tab_layout = QVBoxLayout()
        self.data_tab.setLayout(self.data_tab_layout)

        # 添加文件选择按钮
        self.btn_open = QPushButton("Open Excel File")
        self.btn_open.clicked.connect(self.openFileDialog)  # 绑定按钮点击事件
        self.data_tab_layout.addWidget(self.btn_open)

        # 添加文件路径文本框
        self.line_edit = QLineEdit()
        self.data_tab_layout.addWidget(self.line_edit)

        # 添加表格视图
        self.table_view = QTableView()
        self.data_tab_layout.addWidget(self.table_view)  # 将表格视图加入布局

        self.show()  # 显示窗口

    def openFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
        if fileName:
            self.line_edit.setText(fileName)  # 设置文件路径到文本框
            self.loadExcel(fileName)  # 加载Excel文件

    def loadExcel(self, file_path):
        try:
            self.df = pd.read_excel(file_path)  # 读取Excel文件
            model = PandasModel(self.df)
            self.table_view.setModel(model)  # 设置模型到表格视图
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            # 可以在这里添加一些GUI错误提示

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelViewer()
    sys.exit(app.exec_())
