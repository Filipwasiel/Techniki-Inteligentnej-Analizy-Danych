from PySide6.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit, QTextEdit, QComboBox, \
    QListWidget, QHBoxLayout, QCheckBox, QSizePolicy, QGroupBox, QFormLayout, QSpinBox
from PySide6.QtCore import Qt

objects = [
    'first object',
    'second object',
    'third object',
    'fourth object',
    'fifth object',
]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Converter xlsx to docs/pdf")
        self.resize(800, 600)
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.layout = QVBoxLayout(main_widget)
        self.settings = {}

        self.create_widgets()

    def create_widgets(self):
        self.layout.addWidget(QLabel("<b>1. Choose file: </b>"))
        self.btn_browse = QPushButton("Browse")
        self.btn_browse.clicked.connect(self.browse_file)
        self.layout.addWidget(self.btn_browse)

        self.file_label = QLabel("Brak wybranego pliku")
        self.file_label.setStyleSheet("color: red;")
        self.layout.addWidget(self.file_label)

        self.layout.addWidget(QLabel("<b>2. Please provie title for document: </b>"))
        self.title_entry = QLineEdit()
        self.layout.addWidget(self.title_entry)

        settings_group = QGroupBox("3. User Settings")
        form_layout = QFormLayout()

        self.font_size_input = QLineEdit(self.settings.get("font_size", "11"))
        self.margin_input = QLineEdit(self.settings.get("margin", "2.0"))
        form_layout.addRow("Font size: ", self.font_size_input)
        form_layout.addRow("Margin: ", self.margin_input)

        btn_save = QPushButton("Save settings")
        btn_save.clicked.connect(self.save_user_settings)
        form_layout.addRow(btn_save)

        settings_group.setLayout(form_layout)
        self.layout.addWidget(settings_group)


    def browse_file(self):
        print("wybieranie pliku")
        self.file_label.setText("Wybrano plik!!")
        self.file_label.setStyleSheet("color: green;")

    def save_user_settings(self):
        print("Ustawienia zapisane!!")