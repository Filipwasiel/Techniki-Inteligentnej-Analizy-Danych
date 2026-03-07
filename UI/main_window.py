import os
from Logic.converter import DocumentConverter
from Utils.settings_loader import load_settings, save_settings
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton,
                               QLineEdit, QGroupBox, QFormLayout, QMessageBox, QHBoxLayout, QFileDialog)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Converter xlsx to docs/pdf")
        self.resize(800, 600)

        self.filepath = None
        self.settings = load_settings()

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.layout = QVBoxLayout(main_widget)

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

        self.layout.addWidget(QLabel("<b>4. Generuj:</b>"))
        btn_layout = QHBoxLayout()

        self.btn_docx = QPushButton("Convert to docx")
        self.btn_docx.setStyleSheet("background-color: #add8e6;")
        self.btn_docx.clicked.connect(lambda: self.process_conversion("docx"))

        self.btn_pdf = QPushButton("Konwertuj do PDF")
        self.btn_pdf.setStyleSheet("background-color: #90ee90;")
        self.btn_pdf.clicked.connect(lambda: self.process_conversion("pdf"))

        btn_layout.addWidget(self.btn_docx)
        btn_layout.addWidget(self.btn_pdf)
        self.layout.addLayout(btn_layout)


    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel Files (*.xlsx)")
        if file_path:
            self.filepath = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.file_label.setStyleSheet("color: green;")

    def save_user_settings(self):
        new_settings = {
            "font_size": self.font_size_input.text(),
            "margin": self.margin_input.text(),
        }
        save_settings(new_settings)
        QMessageBox.information(self, 'Zapisano', 'Ustawienia zostały zapisane!')

    def process_conversion(self, format_type):
        if not self.filepath:
            QMessageBox.critical(self, "Błąd", "Najpierw wybierz plik Excel!")
            return

        base_name = os.path.splitext(self.filepath)[0]
        docx_path = f"{base_name}.docx"
        title = self.title_entry.text()
        font_size = self.font_size_input.text()
        margin = self.margin_input.text()

        try:
            DocumentConverter.generate_docx(self.filepath, docx_path, title, font_size, margin)

            if format_type == "docx":
                QMessageBox.information(self, "Sukces", f"Zapisano plik:\n{docx_path}")
            elif format_type == "pdf":
                pdf_path = f"{base_name}.pdf"
                DocumentConverter.generate_pdf(docx_path, pdf_path)
                QMessageBox.information(self, "Sukces", f"Zapisano plik:\n{pdf_path}")

        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Wystąpił błąd podczas konwersji:\n{str(e)}")