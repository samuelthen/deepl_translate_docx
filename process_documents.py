import deepl, fitz, re, os, locale, time, shutil
from docx2pdf import convert
from pypdf import PdfWriter, PdfReader, PdfMerger
from docx.shared import Cm, Pt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from datetime import datetime
import deepl

config = {
    "auth_key" : "", # copy key from DeepL
    "input_folder" : "./input_folder",
    "output_folder" : "./pdf_folder/",
    "temporary_files_path": "./lib/temporary_files/",
    "template_files_path": "./lib/templates/",
    "font_file": "./lib/font/ArialUnicodeBold.ttf",
    "font_name": "ArialUnicodeBold",
    "margin": {"top": 6, "bottom": 1.5, "left": 8, "right": 1.5},  # For IS cover page
    "fontsize": 8.5 # IS cover page
}

def delete_files():
    # Delete temporary files to avoid unexpected results

    folder_path = config["temporary_files_path"]
    max_attempts = 5  # Maximum number of attempts to delete a file
    attempt_delay = 1  # Delay in seconds between attempts
    
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # Check if it's a file or symlink
        if os.path.isfile(file_path) or os.path.islink(file_path):
            attempts = 0
            while attempts < max_attempts:
                try:
                    os.remove(file_path)
                    break  # If we got here, the file was removed successfully
                except PermissionError:
                    time.sleep(attempt_delay)
                except Exception as e:
                    print(f"An error occurred while deleting {file_path}: {e}")
                    break  # Break the loop to avoid an infinite loop
                attempts += 1
            else:
                print(f"Failed to delete {file_path} after {max_attempts} attempts.")

        elif os.path.isdir(file_path):
            try:
                shutil.rmtree(file_path)
            except Exception as e:
                print(f"An error occurred while deleting {file_path}: {e}")

def get_file_info(input_path):
    try:
        filename = input_path.split('/')[-1]
        filename = input_path.split('\\')[-1].split('.')[0]
        # Retrieve company or industry name
        pattern = r'^\d+|\bJA\b|\bZH\b|\bDE\b|\b\d+\s\w+\s\d+\b'   
        company_name = re.sub(pattern, '', filename).strip()  # Can be replaced to extract company name from database
        with open(input_path, 'rb') as f:
            doc = Document(f)
 
        # Extract reference number and determine document id and date
        found_acq_ref = False
        for p in doc.paragraphs:
            if found_acq_ref:
                break

            if 'ACQ_REF' in p.text:
                ref = re.search(r"[\w\d]+\s*(\/\s*[\w\d]+\s*)+", p.text).group(0)
                ref_parts = ref.split('/')
                doc_type = ref_parts[0]
                id = ref_parts[1]
                date = ref_parts[2]
                company_code = ref_parts[3]
                found_acq_ref = True

        if not found_acq_ref:
            print(f"ACQ_REF not found for {filename}")
            return

        filename_parts = filename.split(' ')

        # Detect language specified in file name
        if filename_parts[-1]=="JA" or filename_parts[-1]=="ZH":
            language = filename_parts[-1]
            date = f"{date[:4]}年{date[4:6]}月{date[6:]}日"
        elif filename_parts[-1]=="DE":
            language = filename_parts[-1]
            locale.setlocale(locale.LC_TIME, 'de_DE.UTF-8')
            date_obj = datetime.strptime(date, "%Y%m%d")
            date = date_obj.strftime("%d %B %Y")
        else:
            language = "EN"
            locale.setlocale(locale.LC_TIME, 'en.UTF-8')
            date_obj = datetime.strptime(date, "%Y%m%d")
            date = date_obj.strftime("%d %B %Y")

        file_info = {
           "filename": filename,
           "company_name": company_name,
           "company_code": company_code,
           "doc_type": doc_type,
           "id": id,
           "date": date,
           "language": language
        }

        return file_info
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while getting file info.")
        return None


'''
# Alternative way to extract company name
from pymongo import MongoClient
def extract_company_name(company_code):

    try:
                # replace xxxxx with database ip address 
        client = MongoClient("mongodb://xxxxx:27017/") 
        # replace xxxxx with database name 
        db = client['xxxxx'] 
        # replace xxxxx with table wanted to used
        col = db['xxxxx'] 
        print("Connect to DB.")
    except:
        print("Can't connect to DB.")    

    documents = col.find({"compnumber": 111754})

    for document in documents:
        print(document['name'])

'''

def delete_acq_ref(input_path):
    try:
        with open(input_path, 'rb') as f:
            doc = Document(f)

        found_acq_ref = False
        for p in doc.paragraphs:
            if found_acq_ref:
                break

            if 'ACQ_REF' in p.text:
                found_acq_ref = True
                p.text = ""

        found_acq_author = False
        for p in doc.paragraphs:
            if found_acq_author:
                break

            if 'ACQ_AUTHOR' in p.text:
                p.text = ""
                found_acq_author = True

        file_path = os.path.join(config['temporary_files_path'], "input.docx")
        with open(file_path, 'wb') as f:
            doc.save(f)
        
        return file_path
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while deleting ACQ_REF.")
        return None

def create_CS_pdf(docx_file, doc_type, language):
    try:
        content =os.path.join(config['temporary_files_path'], "content.pdf")
        convert(docx_file, content)
        
        path = os.path.join(config['template_files_path'], f"cover-{doc_type}-{language}.pdf")
        cover_page = fitz.open(path)
        content = fitz.open(content)

        cover_page.insert_file(content)

        merged_pdf = os.path.join(config['temporary_files_path'], "output.pdf")
        cover_page.save(merged_pdf)

        return merged_pdf
    
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while creating the CS PDF.")
        return None

def edit_CS_pdf(merged_pdf, company_name, date, id, language, filename,output_path):
    try:
        doc = fitz.open(merged_pdf)
        number = f"NO.: {id}-{language}"
        page = doc[0]
        rect = fitz.Rect(90, 265, 650, 400)
        page.insert_textbox(rect, company_name, fontsize=28, fontname=config["font_name"],
                                 fontfile=config["font_file"], fill=(1, 0, 0), align=0)

        rect1 = fitz.Rect(90, 320, 650, 400)
        page.insert_textbox(rect1, date, fontsize=12, fontname=config["font_name"],
                                  fontfile=config["font_file"], align=0)

        rect2 = fitz.Rect(90, 110, 600, 400)
        page.insert_textbox(rect2, number, fontsize=9, fontname=config["font_name"],
                                  fontfile=config["font_file"], align=0)

        doc.save(os.path.join(output_path, f"{filename}.pdf"))
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while editing the CS PDF.")

def extract_IS_cover_page(input_path):
    try:
        with open(input_path, 'rb') as f:
            doc = Document(f)
        for table in doc.tables:
            table._element.getparent().remove(table._element)
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.startswith("#(1)"):
                doc.paragraphs[i]._element.getparent().remove(doc.paragraphs[i]._element)
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.startswith("#(2)"):
                while len(doc.paragraphs) > i:
                    doc.paragraphs[i]._element.getparent().remove(doc.paragraphs[i]._element)

        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(config["margin"]["top"])
            section.bottom_margin = Cm(config["margin"]["bottom"])
            section.left_margin = Cm(config["margin"]["left"])
            section.right_margin = Cm(config["margin"]["right"])

        for paragraph in doc.paragraphs:
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in paragraph.runs:
                font = run.font
                font.size = Pt(config['fontsize'])

        for section in doc.sections:
            header = section.header
            if header is not None:
                for p in header.paragraphs:
                    p.text = ''
            footer = section.footer
            if footer is not None:
                for p in footer.paragraphs:
                    p.text = ''

        path = os.path.join(config['temporary_files_path'], "cover_page.docx")
        with open(path, 'wb') as f:
            doc.save(f)
        
        cover_page = os.path.join(config['temporary_files_path'], "cover_page.pdf")
        convert(path, cover_page)

        return cover_page
    
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while extracting IS cover page.")
        return None
    
def extract_IS_content(input_path):
    try:
        with open(input_path, 'rb') as f:
            doc = Document(f)
        for i, para in enumerate(doc.paragraphs):
            if para.text.startswith("#(3)"):
                break
        for _ in range(i):
            doc.paragraphs[0]._element.getparent().remove(doc.paragraphs[0]._element)
        doc.paragraphs[0]._element.getparent().remove(doc.paragraphs[0]._element)

        path = os.path.join(config['temporary_files_path'],"content.docx")
        with open(path, 'wb') as f:
            doc.save(f)
        
        content = f"{config['temporary_files_path']}/content.pdf"
        convert(path, content)
        return content
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while extracting IS content.")
        return None

def create_IS_cover_page(cover_page, doc_type, language):
    try:
        template = PdfReader(open(os.path.join(config['template_files_path'],f"cover-{doc_type}-{language}.pdf"), "rb"))
        cover_content = PdfReader(open(cover_page, "rb"))
        for i in range(len(cover_content.pages)):
            background = template.pages[i]
            foreground = cover_content.pages[i]
            background.merge_page(foreground)

        writer = PdfWriter()
        number_of_cover = len(cover_content.pages)
        for i in range(number_of_cover):
            writer.add_page(template.pages[i])

        with open(cover_page, "wb") as outFile:
            writer.write(outFile)
        return number_of_cover
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while creating IS cover page.")
        return None

def merge_IS_pdf(cover_page, content):
    try:
        merger = PdfMerger()
        merger.append(cover_page)
        merger.append(content)

        merged_pdf = os.path.join(config['temporary_files_path'], "output.pdf")
        with open(merged_pdf, "wb") as outFile:
            merger.write(outFile)
        merger.close()

        return merged_pdf
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while merging IS PDFs.")
        return None

def edit_IS_pdf(merged_pdf, company_name, date, id, language, number_of_cover, filename, output_path):
    try:
        doc = fitz.open(merged_pdf)
        number = f"NO.: {id}-{language}"
        page = doc[0]
        
        if language != "EN":
            translator = deepl.Translator(config["auth_key"])
            result = translator.translate_text(company_name, target_lang=language)
            company_name = result.text

        def add_text(page):
            rect = fitz.Rect(225, 100, 650, 400)
            page.insert_textbox(rect, company_name, fontsize=16, fontname=config["font_name"],
                                    fontfile=config["font_file"], fill=(1, 0, 0), align=0)

            rect1 = fitz.Rect(225, 130, 650, 400)
            page.insert_textbox(rect1, date, fontsize=9, fontname=config["font_name"],
                                    fontfile=config["font_file"], align=0)

            rect2 = fitz.Rect(120, 80, 600, 400)
            page.insert_textbox(rect2, number, fontsize=9, fontname=config["font_name"],
                                    fontfile=config["font_file"], align=0)
        add_text(page)

        for i in range(1, number_of_cover):  # Assuming the information is added to all the cover pages
            page = doc[i]
            add_text(page)

        doc.save(os.path.join(output_path, f"{filename}.pdf"))
    except Exception as e:
        print(f"{str(e)}\nAn error occurred while editing IS PDF.")

def main():
    translated_files = []

    for file in os.listdir(config["input_folder"]):
        translated_files.append(os.path.join(config["input_folder"], file))
    for file in translated_files:

        print(file)
        file_info = get_file_info(file)
        filename = file_info["filename"]
        company_name = file_info["company_name"]
        doc_type = file_info["doc_type"]
        id = file_info["id"]
        date = file_info["date"]
        language = file_info["language"]

        if doc_type == "IS":
            file_path = delete_acq_ref(file)
            cover_page = extract_IS_cover_page(file_path)
            content = extract_IS_content(file_path)
            number_of_cover = create_IS_cover_page(cover_page, doc_type, language) 
            merged_pdf = merge_IS_pdf(cover_page, content)
            edit_IS_pdf(merged_pdf, company_name, date, id, language, number_of_cover, filename, config["output_folder"])
           
        elif doc_type == "CS":
            file_path = delete_acq_ref(file)
            merged_pdf = create_CS_pdf(file_path, doc_type, language)
            edit_CS_pdf(merged_pdf, company_name, date, id, language, filename,config["output_folder"])

        delete_files()


if __name__ == "__main__":
    main()