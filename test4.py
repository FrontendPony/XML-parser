import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

def main():
    app = QApplication(sys.argv)
    window = QMainWindow()
    tableWidget = QTableWidget()
    window.setCentralWidget(tableWidget)

    tableWidget.setRowCount(5)  # Set the number of rows
    tableWidget.setColumnCount(5)  # Set the number of columns

    # Set data in a specific cell (e.g., row 2, column 3)
    cell_value = "Hello, World!"
    item = QTableWidgetItem(cell_value)
    tableWidget.setItem(0, 0, item)

    # Color the background of the specific cell (e.g., row 2, column 3) with a red color
    item.setBackground(QColor(255, 0, 0))

    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
