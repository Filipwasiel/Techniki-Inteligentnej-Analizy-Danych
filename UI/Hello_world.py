from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QWidget, QVBoxLayout
from PySide6.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hello World")
        self.setFixedSize(400, 300)

        label = QLabel("Hello World")
        label.setAlignment(Qt.AlignCenter)

        self.setCentralWidget(label)
        self.show()