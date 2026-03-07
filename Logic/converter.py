import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert


class DocumentConverter:
    @staticmethod
    def generate_docx(excel_path, output_path, title, font_size, margin_cm):
        df = pd.read_excel(excel_path)
        df = df.fillna("")
        doc = Document()

        try:
            margin = Cm(float(margin_cm))
        except ValueError:
            margin = Cm(2.0)

        for section in doc.sections:
            section.top_margin = margin
            section.left_margin = margin
            section.bottom_margin = margin
            section.right_margin = margin

        if title:
            heading = doc.add_heading(title, 1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        table = doc.add_table(rows=1, cols=len(df.columns))
        table.style = 'Table Grid'

        hdr_cells = table.rows[0].cells
        for i, col_name in enumerate(df.columns):
            hdr_cells[i].text = str(col_name)
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    try:
                        run.font.size = Pt(int(font_size))
                    except ValueError: pass

        for index, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
                for paragraph in row_cells[i].paragraphs:
                    for run in paragraph.runs:
                        try:
                            run.font.size = Pt(int(font_size))
                        except ValueError: pass

        doc.save(output_path)

    @staticmethod
    def generate_pdf(docx_path, output_path):
        convert(docx_path, output_path)
