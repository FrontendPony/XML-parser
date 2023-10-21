import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QWidget

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("TableWidget Example")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.table_widget = QTableWidget(2, 2)
        self.layout.addWidget(self.table_widget)

        self.button = QPushButton("Clear Data and Next")
        self.layout.addWidget(self.button)

        self.button.clicked.connect(self.clear_data_and_next)

        self.current_value = 1
        self.populate_table(1)

    def populate_table(self, value):
        for row in range(2):
            for col in range(2):
                item = QTableWidgetItem(f"Row {row + 1}, Col {col + 1}, Value {value}")
                self.table_widget.setItem(row, col, item)

    def clear_data_and_next(self):
        self.current_value += 1
        if self.current_value > 3:
            self.current_value = 1
        self.table_widget.clearContents()
        self.populate_table(self.current_value)

def main():
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
