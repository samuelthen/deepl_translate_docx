from src.deepl_translate import translate_docx
from src.create_IS_pdf import get_file_info, extract_IS_cover_page, extract_IS_content, create_IS_cover_page,delete_acq_ref, merge_IS_pdf, edit_IS_pdf
from src.create_xml import create_xml
from src.create_CS_pdf import create_CS_pdf, edit_CS_pdf
from src.upload_folder import input_folder
import os

auth_key = "bff9578c-9049-757d-7d43-94017510b368:fx"  # Replace with your DeepL API key
input_path = input_folder("./upload_folder")
target_lang = [["DE", False, False],    # DE for German, JA for Japanese and ZH for Chinese
               ["JA", False, False],   # [language code, translate flag, use of glossary flag]
               ["ZH", False, False]
               ]
output_folder = "./output_folder"
glossary_data_path = "./lib/glossary_data"
temporary_files_path = "./lib/temporary_files"
template_folder_path = "./lib/templates"
fontfile = "./lib/font/ArialUnicodeBold.ttf"

original_filename = input_path.split('/')[-1].split('.')[0]
output_folder_path = f"./{output_folder}/{original_filename}"
if os.path.exists(f"./{output_folder}/{original_filename}") == True:
    pass
else:
    os.makedirs(f"./{output_folder}/{original_filename}")

translated_files = translate_docx(auth_key, input_path, target_lang, output_folder_path, glossary_data_path)

for file in translated_files:
    
    file_info = get_file_info(file)
    filename = file_info["filename"]
    company_name = file_info["company_name"]
    doc_type = file_info["doc_type"]
    id = file_info["id"]
    date = file_info["date"]
    language = file_info["language"]

    if doc_type == "IS":
        file = delete_acq_ref(file,temporary_files_path)
        cover_page = extract_IS_cover_page(file,temporary_files_path)
        content = extract_IS_content(file,temporary_files_path)
        number_of_cover = create_IS_cover_page(cover_page, template_folder_path, doc_type, language) 
        merged_pdf = merge_IS_pdf(cover_page, content, temporary_files_path)
        edit_IS_pdf(merged_pdf, company_name, date, id, language, number_of_cover, filename, output_folder_path, fontfile)
    elif doc_type == "CS":
        file = delete_acq_ref(file,temporary_files_path)
        merged_pdf = create_CS_pdf(file, template_folder_path, temporary_files_path, doc_type, language)
        edit_CS_pdf(merged_pdf, company_name, date, id, language, filename, output_folder_path,fontfile)
    
