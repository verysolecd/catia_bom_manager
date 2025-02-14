import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout
import openpyxl

# 主窗口模块


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 创建按钮
        self.button = QPushButton('打开 Excel 并注入数据', self)
        self.button.clicked.connect(self.open_excel_and_inject_data)

        # 布局
        layout = QVBoxLayout()
        layout.addWidget(self.button)
        self.setLayout(layout)

        self.setWindowTitle('Excel 操作示例')
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def open_excel_and_inject_data(self):
        # 调用 Excel 操作模块的函数
        excel_operator = ExcelOperator()
        excel_operator.open_excel()
        excel_operator.inject_data()
        # 提示用户手动操作 Excel 并保存
        input("请手动打开 Excel 文件进行操作，操作完成后保存并按回车键继续...")
        data = excel_operator.read_data_to_array()
        # 调用数据处理模块的函数
        data_processor = DataProcessor()
        data_processor.process_data(data)

# Excel 操作模块


class ExcelOperator:
    def __init__(self):
        self.workbook = None
        self.worksheet = None

    def open_excel(self):
        # 打开或创建 Excel 文件
        try:
            self.workbook = openpyxl.load_workbook('example.xlsx')
        except FileNotFoundError:
            self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active

    def inject_data(self):
        # 向 Excel 中注入数据
        data_to_inject = [
            ['姓名', '年龄', '性别'],
            ['张三', 25, '男'],
            ['李四', 30, '女']
        ]
        for row in data_to_inject:
            self.worksheet.append(row)
        self.workbook.save('example.xlsx')

    def read_data_to_array(self):
        # 读取 Excel 数据到数组
        data = []
        for row in self.worksheet.iter_rows(values_only=True):
            data.append(list(row))
        return data

# 数据处理模块


class DataProcessor:
    def process_data(self, data):
        # 对从 Excel 中提取的数据进行处理
        print("从 Excel 中提取的数据：")
        for row in data:
            print(row)
        # 可以在这里添加更多的数据处理逻辑，例如提取特定列的数据到数组

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())
