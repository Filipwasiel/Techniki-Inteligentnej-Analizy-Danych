import re
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert


class DocumentConverter:
    @staticmethod
    def clean_text(text):
        text = str(text)
        return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)

    @staticmethod
    def add_page_number(run):
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    @staticmethod
    def generate_docx(excel_path, output_path, title, config):
        df = pd.read_excel(excel_path)
        if "L.p." in df.columns:
            df = df.drop(columns=["L.p."])
        df = df.fillna("")
        doc = Document()

        font_name = config.get("font_family", "Arial")
        font_size = int(config.get("font_size", 12))
        line_spacing = float(config.get("line_spacing", 1.15))
        margin = Cm(float(str(config.get("margin", 2.0)).replace(',', '.')))

        for section in doc.sections:
            section.top_margin = section.bottom_margin = margin
            section.left_margin = section.right_margin = margin
            if config.get("page_numbers", True):
                footer = section.footer
                footer_p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                footer_p.add_run("Strona ")
                DocumentConverter.add_page_number(footer_p.add_run())

        if title:
            heading = doc.add_heading(title, 0)
            # heading.name = font_name
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if config.get("format_type") == "Table":
            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = 'Light Shading'
            hdr_cells = table.rows[0].cells
            for i, column in enumerate(df.columns):
                hdr_cells[i].text = str(column)

            for _, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, val in enumerate(row):
                    row_cells[i].text = DocumentConverter.clean_text(val)
        else:
            for index, row in df.iterrows():
                separator = doc.add_paragraph()
                separator.paragraph_format.keep_with_next = True
                separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_sep = separator.add_run(f"Rekord {index + 1}")
                run_sep.bold = True
                run_sep.name = font_name
                run_sep.font.size = Pt(font_size + 2)
                items = list(row.items())
                for i, (name, val) in enumerate(items):
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = line_spacing
                    if i < len(items) - 1:
                        p.paragraph_format.keep_with_next = True
                    p.paragraph_format.keep_together = True
                    run_col = p.add_run(f"{DocumentConverter.clean_text(name)}: ")
                    run_col.bold = True
                    run_col.font.name = font_name
                    run_col.font.size = Pt(font_size)
                    run_val = p.add_run(str(DocumentConverter.clean_text(val)))
                    run_val.font.name = font_name
                    run_val.font.size = Pt(font_size)

        doc.save(output_path)

    @staticmethod
    def generate_pdf(docx_path, output_path):
        convert(docx_path, output_path)
