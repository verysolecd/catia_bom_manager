import sys
import csv
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QFileDialog

class TableApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 创建布局
        layout = QVBoxLayout()

        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(11)  # 设置列数
        # 定义表头
        headers = '''
            '<html><head/><body><p>质量<br>Mass</p></body></html>',
            '<html><head/><body><p>厚度<br>Thickness</p></body></html>',
            '<html><head/><body><p>零件号<br>Partnumber</p></body></html>',
            '<html><head/><body><p>更改<br>件号</p></body></html>',
            '<html><head/><body><p>英文名称<br>Nomenclature</p></body></html>',
            '<html><head/><body><p>更改<br>英文名</p></body></html>',
            '<html><head/><body><p>中文名称<br>Definition</p></body></html>',
            '<html><head/><body><p>更改<br>中文名</p></body></html>',
            '<html><head/><body><p>实例名<br>InstanceName</p></body></html>',
            '<html><head/><body><p>更改<br>实例名</p></body></html>',
            '<html><head/><body><p>材料<br>material</p></body></html>',
            '<html><head/><body><p>定义<br>材料</p></body></html>',
            '<html><head/><body><p>密度<br>Material</p></body></html>',
            '<html><head/><body><p>更改<br>密度</p></body></html>',

        '''.splitlines()

        # 设置列标题
        self.table.setHorizontalHeaderLabels(headers)
        layout.addWidget(self.table)

        # 创建读取按钮
        read_button = QPushButton('读取数据')
        read_button.clicked.connect(self.read_data)
        layout.addWidget(read_button)

        # 创建保存按钮
        save_button = QPushButton('修改数据')
        save_button.clicked.connect(self.save_data)
        layout.addWidget(save_button)

        # 设置布局
        self.setLayout(layout)
        self.setWindowTitle('Table Data Read/Write')
        self.show()

    def read_data(self):
        # 打开文件选择对话框
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open CSV File', '', 'CSV Files (*.csv)')
        if file_path:
            with open(file_path, 'r', newline='') as file:
                reader = csv.reader(file)
                data = list(reader)

            # 清空表格
            self.table.setRowCount(0)

            # 填充表格数据
            for row_index, row_data in enumerate(data):
                self.table.insertRow(row_index)
                for col_index, cell_data in enumerate(row_data):
                    item = QTableWidgetItem(cell_data)
                    self.table.setItem(row_index, col_index, item)

    def save_data(self):
        # 打开文件保存对话框
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save CSV File', '', 'CSV Files (*.csv)')
        if file_path:
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)
                rows = self.table.rowCount()
                cols = self.table.columnCount()
                for row in range(rows):
                    row_data = []
                    for col in range(cols):
                        item = self.table.item(row, col)
                        if item is not None:
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    writer.writerow(row_data)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TableApp()
    sys.exit(app.exec_())