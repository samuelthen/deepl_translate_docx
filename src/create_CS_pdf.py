from docx2pdf import convert
from pypdf import PdfWriter, PdfReader, PdfMerger
import fitz
from docx.shared import Cm, Pt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
import re

def create_CS_pdf(file, template_files_path,temporary_files_path, doc_type, language):
    content = f"{temporary_files_path}/content.pdf"
    convert(file, content)
    cover_page = PdfReader(open(f"{template_files_path}/cover-{doc_type}-{language}.pdf", "rb"))
    content = PdfReader(open(content, "rb"))
    
    pdfs = [cover_page, content]
    merger = PdfMerger()
    for pdf in pdfs:
        merger.append(pdf)
    merged_pdf = f"{temporary_files_path}/output.pdf"
    merger.write(merged_pdf)
    merger.close()
    return merged_pdf

def edit_CS_pdf(merged_pdf, company_name, date, id, language, filename, output_folder_path,fontfile):
    doc = fitz.open(merged_pdf) 
    number = f"NO. {id}-{language}"
    page = doc[0]
    rect = fitz.Rect(215, 285, 650, 400) 
    rc = page.insert_textbox(rect, company_name, fontsize = 16, # choose fontsize (float)
                        fontname = "ArialUnicodeBold",       # a PDF standard font
                        fontfile = fontfile,
                        fill=(1, 0, 0),              
                        align = 0)                      # 0 = left, 1 = center, 2 = right

    rect1 = fitz.Rect(215, 330, 650, 400) 
    rc1 = page.insert_textbox(rect1, date, fontsize = 9, # choose fontsize (float)
                        fontname = "ArialUnicodeBold",       # a PDF standard font
                        fontfile = fontfile,
                        align = 0)                      # 0 = left, 1 = center, 2 = right

    rect2 = fitz.Rect(130, 110, 600, 400)
    rc1 = page.insert_textbox(rect2, number, fontsize = 9, # choose fontsize (float)
                        fontname = "ArialUnicodeBold",       # a PDF standard font
                        fontfile = fontfile,
                        align = 0)                      # 0 = left, 1 = center, 2 = right

    doc.save(f"{output_folder_path}/{filename}.pdf")   # Update file by doc.saveIncr(). Save to new instead by doc.save("new.pdf",...)