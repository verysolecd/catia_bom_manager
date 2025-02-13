import sys
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QMenu, QAction
from PyQt5.QtGui import QClipboard


class TableWidgetWithContextMenu(QTableWidget):
    def __init__(self, rows, columns, parent=None):
        super().__init__(rows, columns, parent)

    def contextMenuEvent(self, event):
        # 创建右键菜单
        menu = QMenu(self)

        # 创建菜单项
        copy_action = QAction("复制", self)
        paste_action = QAction("粘贴", self)
        delete_action = QAction("删除", self)

        # 为菜单项添加触发事件
        copy_action.triggered.connect(self.copy_cells)
        paste_action.triggered.connect(self.paste_cells)
        delete_action.triggered.connect(self.delete_cells)

        # 将菜单项添加到菜单中
        menu.addAction(copy_action)
        menu.addAction(paste_action)
        menu.addAction(delete_action)

        # 在鼠标点击位置显示菜单
        menu.exec_(event.globalPos())

    def copy_cells(self):
        selected_ranges = self.selectedRanges()
        if selected_ranges:
            copy_text = ""
            for selected_range in selected_ranges:
                for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                    for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                        item = self.item(row, col)
                        if item:
                            copy_text += item.text()
                        if col < selected_range.rightColumn():
                            copy_text += "\t"
                    copy_text += "\n"
            clipboard = QApplication.clipboard()
            clipboard.setText(copy_text)

    def paste_cells(self):
        clipboard = QApplication.clipboard()
        paste_text = clipboard.text()
        rows = paste_text.split('\n')
        selected_ranges = self.selectedRanges()
        if selected_ranges:
            top_left_row = selected_ranges[0].topRow()
            top_left_col = selected_ranges[0].leftColumn()
            for i, row in enumerate(rows):
                cols = row.split('\t')
                for j, col in enumerate(cols):
                    target_row = top_left_row + i
                    target_col = top_left_col + j
                    if target_row < self.rowCount() and target_col < self.columnCount():
                        item = QTableWidgetItem(col)
                        self.setItem(target_row, target_col, item)

    def delete_cells(self):
        selected_ranges = self.selectedRanges()
        for selected_range in selected_ranges:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    self.setItem(row, col, QTableWidgetItem())


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()

        # 创建带有右键菜单的表格
        self.table = TableWidgetWithContextMenu(5, 5)
        for row in range(5):
            for col in range(5):
                item = QTableWidgetItem(f"Cell {row}-{col}")
                self.table.setItem(row, col, item)

        layout.addWidget(self.table)
        self.setLayout(layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
