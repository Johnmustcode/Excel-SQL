import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog, QTableView, QWidget
from PyQt5.QtWidgets import QComboBox
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

        # 创建中央布局
        central_widget = QWidget()
        central_layout = QVBoxLayout()
        central_widget.setLayout(central_layout)
        self.setCentralWidget(central_widget)

        # 添加文件选择按钮
        self.btn_open = QPushButton("Open Excel File")
        self.btn_open.clicked.connect(self.openFileDialog)  # 绑定按钮点击事件
        central_layout.addWidget(self.btn_open)
        
        # 添加文件路径文本框
        self.line_edit = QLineEdit()
        central_layout.addWidget(self.line_edit)

        # 创建工作表的标签页
        self.tabs = QTabWidget()
        central_layout.addWidget(self.tabs)

        

        self.show()  # 显示窗口

    def openFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
        if fileName:
            self.line_edit.setText(fileName)  # 设置文件路径到文本框
            self.loadExcelFile(fileName)  # 加载Excel文件

    def loadExcel(self, file_path):
        try:
            self.df = pd.read_excel(file_path)  # 读取Excel文件
            model = PandasModel(self.df)
            self.table_view.setModel(model)  # 设置模型到表格视图
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            # 可以在这里添加一些GUI错误提示

    def loadExcelFile(self, file_path):
        try:
            # 读取所有工作表
            with pd.ExcelFile(file_path) as xls:
                for sheet_name in xls.sheet_names:
                    # 为每个工作表创建一个新的标签页
                    self.createSheetTab(sheet_name, file_path)
        except Exception as e:
            print(f"Error reading Excel file: {e}")
    
    def createSheetTab(self, sheet_name, file_path):
        # 创建新的QWidget作为标签页内容
        sheet_tab = QWidget()
        sheet_tab_layout = QVBoxLayout()
        sheet_tab.setLayout(sheet_tab_layout)

        # 为这个工作表创建一个QTableView
        sheet_table_view = QTableView()
        sheet_tab_layout.addWidget(sheet_table_view)

        # 读取工作表数据并设置到QTableView
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        model = PandasModel(df)
        sheet_table_view.setModel(model)

        # 将创建的标签页添加到主标签页
        self.tabs.addTab(sheet_tab, sheet_name)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelViewer()
    sys.exit(app.exec_())
