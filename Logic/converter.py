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
        try:
            margin = Cm(float(config["margin"]))
        except ValueError:
            margin = Cm(2.0)

        alignments_map = {
            "Do lewej": WD_ALIGN_PARAGRAPH.LEFT,
            "Do prawej": WD_ALIGN_PARAGRAPH.RIGHT,
            "Wyjustuj": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "Do środka": WD_ALIGN_PARAGRAPH.CENTER,
        }

        for section in doc.sections:
            section.top_margin = margin
            section.left_margin = margin
            section.bottom_margin = margin
            section.right_margin = margin
            if config.get("page_numbers", True):
                footer = section.footer
                footer_p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                footer_p.text = ""
                footer_p.add_run("Strona ")
                DocumentConverter.add_page_number(footer_p.add_run())

        if title:
            heading = doc.add_heading(title, 0)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for index, row in df.iterrows():
            separator = doc.add_paragraph()
            separator.paragraph_format.keep_with_next = True
            separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_sep = separator.add_run(f"Rekord {index + 1}")
            run_sep.bold = True
            run_sep.font.size = Pt(int(config["font_size"]) + 2)
            items = list(row.items())
            for i, (name, val) in enumerate(items):
                p = doc.add_paragraph()

                if i < len(items) - 1:
                    p.paragraph_format.keep_with_next = True
                p.paragraph_format.keep_together = True
                run_col = p.add_run(f"{DocumentConverter.clean_text(name)}: ")
                run_col.bold = True
                run_col.font.size = Pt(int(config["font_size"]))
                run_val = p.add_run(str(DocumentConverter.clean_text(val)))
                run_val.font.size = Pt(int(config["font_size"]))

        # doc.add_page_break()
        doc.save(output_path)

    @staticmethod
    def generate_pdf(docx_path, output_path):
        convert(docx_path, output_path)
