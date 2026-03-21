import os
from Logic.converter import DocumentConverter
from Utils.settings_loader import load_settings, save_settings
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton,
                               QLineEdit, QGroupBox, QFormLayout, QMessageBox, QHBoxLayout, QFileDialog, QCheckBox, QApplication,
                               QComboBox, QDoubleSpinBox)
from PySide6.QtGui import Qt, QCursor


class MainWindow(QMainWindow):
   def __init__(self):
      super().__init__()
      self.setWindowTitle("Converter xlsx to docs/pdf")
      self.resize(500, 650)
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

      settings_group = QGroupBox("3. Document Settings")
      form_layout = QFormLayout()
      # Format wyświetlania danych
      self.combo_format = QComboBox()
      self.combo_format.addItems(["List", "Table"])
      self.combo_format.setCurrentText(self.settings.get("format_type", "List"))
      #Ustawienia dla stylów tabeli
      self.combo_table_style = QComboBox()
      self.combo_table_style.addItems(["Light Grid", "Table Grid", "Medium Shading 1 Accent 1"])
      self.combo_table_style.setCurrentText(self.settings.get("table_style", "Light Grid"))
      # W zależności co jest wybrane jako format powyzsze opcje beda wyswietlane lub nie
      self.combo_format.currentTextChanged.connect(self.toggle_table_style)
      self.toggle_table_style(self.combo_format.currentText())
      # Ustawienia związane z orientacją strony
      self.combo_orientation = QComboBox()
      self.combo_orientation.addItems(["Vertical", "Horizontal"])
      self.combo_orientation.setCurrentText(self.settings.get("orientation", "Vertical"))
      # Ustawienia związane z czcionką
      self.font_size_input = QLineEdit(self.settings.get("font_size", "11"))
      self.combo_font_name = QComboBox()
      self.combo_font_name.addItems(["Arial", "Calibri", "Times New Roman", "Verdana"])
      self.combo_font_name.setCurrentText(self.settings.get("font_family", "Arial"))
      #Ustawienia związane z marginesami
      self.margin_val = QDoubleSpinBox()
      self.margin_val.setRange(0.5, 5.0)
      self.margin_val.setValue(float(self.settings.get("margin", 2.0)))
      #Ustawienia związane linespacing
      self.line_spacing = QDoubleSpinBox()
      self.line_spacing.setRange(1.0, 3.0)
      self.line_spacing.setSingleStep(0.1)
      self.line_spacing.setValue(float(self.settings.get("line_spacing", 1.15)))
      #Ustawienia związane z numerowaniem stron
      self.page_num_check = QCheckBox("Page numeration")
      self.page_num_check.setChecked(self.settings.get("page_numbers", True))
      #Ustawienia związane z wyrównaniem tekstu
      self.combo_alignment = QComboBox()
      self.combo_alignment.addItems(["To Left", "Centered", "To Right", "Justify"])
      self.combo_alignment.setCurrentText(self.settings.get("alignment", "To Left"))
      #Style nagłówków w tabeli lub klczy w liście
      self.check_bold = QCheckBox("Bold headers or keys")
      self.check_bold.setChecked(self.settings.get("bold_headers", True))
      self.check_italic = QCheckBox("Italic headers or keys")
      self.check_italic.setChecked(self.settings.get("italic_headers", False))

      #Formularz
      form_layout.addRow("Display format: ", self.combo_format)
      form_layout.addRow("Table style:", self.combo_table_style)
      form_layout.addRow("Orientation: ", self.combo_orientation)
      form_layout.addRow("Text alignment: ", self.combo_alignment)
      form_layout.addRow("Font family: ", self.combo_font_name)
      form_layout.addRow("Font size: ", self.font_size_input)
      form_layout.addRow("Line spacing: ", self.line_spacing)
      form_layout.addRow("Margins (cm): ", self.margin_val)
      form_layout.addRow(self.page_num_check)
      form_layout.addRow(self.check_bold)
      form_layout.addRow(self.check_italic)

      btn_save = QPushButton("Save settings")
      btn_save.clicked.connect(self.save_user_settings)
      form_layout.addRow(btn_save)

      settings_group.setLayout(form_layout)
      self.layout.addWidget(settings_group)

      # self.layout.addWidget(QLabel("<b>4. Generuj:</b>"))
      #Przyciski generowania
      btn_layout = QHBoxLayout()
      self.btn_docx = QPushButton("Generate docx")
      self.btn_docx.setStyleSheet("background-color: #add8e6;")
      self.btn_docx.clicked.connect(lambda: self.process_conversion("docx"))
      self.btn_pdf = QPushButton("Generate PDF")
      self.btn_pdf.setStyleSheet("background-color: #90ee90;")
      self.btn_pdf.clicked.connect(lambda: self.process_conversion("pdf"))
      btn_layout.addWidget(self.btn_docx)
      btn_layout.addWidget(self.btn_pdf)
      self.layout.addLayout(btn_layout)

   def toggle_table_style(self, format_type):
      if format_type == "Table":
         self.combo_table_style.setEnabled(True)
      else:
         self.combo_table_style.setEnabled(False)

   def browse_file(self):
      file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "Utils", "Excel Files (*.xlsx)")
      if file_path:
         self.filepath = file_path
         self.file_label.setText(os.path.basename(file_path))
         self.file_label.setStyleSheet("color: green;")

   def save_user_settings(self):
      new_settings = {
         "format_type": self.combo_format.currentText(),
         "table_style": self.combo_table_style.currentText(),
         "orientation": self.combo_orientation.currentText(),
         "alignment": self.combo_alignment.currentText(),
         "font_family": self.combo_font_name.currentText(),
         "font_size": self.font_size_input.text(),
         "line_spacing": str(self.line_spacing.value()),
         "margin": str(self.margin_val.value()),
         "page_numbers": self.page_num_check.isChecked(),
         "bold_headers": self.check_bold.isChecked(),
         "italic_headers": self.check_italic.isChecked(),
      }
      save_settings(new_settings)
      QMessageBox.information(self, 'Saved', 'Settings were successfully saved!')

   def process_conversion(self, format_type):
      if not self.filepath:
         QMessageBox.critical(self, "Error", "Choose Excel File!")
         return

      self.setEnabled(False)
      QApplication.setOverrideCursor(Qt.WaitCursor)
      QApplication.processEvents()

      base_name = os.path.splitext(self.filepath)[0]
      docx_path = f"{base_name}.docx"
      title = self.title_entry.text()
      config = {
         "format_type": self.combo_format.currentText(),
         "table_style": self.combo_table_style.currentText(),
         "orientation": self.combo_orientation.currentText(),
         "alignment": self.combo_alignment.currentText(),
         "font_family": self.combo_font_name.currentText(),
         "font_size": self.font_size_input.text(),
         "line_spacing": self.line_spacing.value(),
         "margin": self.margin_val.text(),
         "page_numbers": self.page_num_check.isChecked(),
         "bold_headers": self.check_bold.isChecked(),
         "italic_headers": self.check_italic.isChecked(),
      }
      try:
         DocumentConverter.generate_docx(self.filepath, docx_path, title, config)
         if format_type == "docx":
               QMessageBox.information(self, "Sukces", f"Zapisano plik:\n{docx_path}")
         elif format_type == "pdf":
               pdf_path = f"{base_name}.pdf"
               DocumentConverter.generate_pdf(docx_path, pdf_path)
               QMessageBox.information(self, "Sukces", f"Zapisano plik:\n{pdf_path}")

      except Exception as e:
         QMessageBox.critical(self, "Błąd", f"Wystąpił błąd podczas konwersji:\n{str(e)}")
      finally:
         QApplication.restoreOverrideCursor()
         self.setEnabled(True)
