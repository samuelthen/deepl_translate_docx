import deepl 
from docx import Document
import shutil

def translate_without_glossary(auth_key, target_lang, input_path, output_path):
    translator = deepl.Translator(auth_key)
    try:
        translator.translate_document_from_filepath(
            input_path,
            output_path,
            source_lang= "EN",
            target_lang= target_lang
        )
    except deepl.DocumentTranslationException as error:
        # If an error occurs during document translation after the document was
        # already uploaded, a DocumentTranslationException is raised. The
        # document_handle property contains the document handle that may be used to
        # later retrieve the document from the server, or contact DeepL support.
        doc_id = error.document_handle.id
        doc_key = error.document_handle.key
        print(f"Error after uploading ${error}, id: ${doc_id} key: ${doc_key}")
    except deepl.DeepLException as error:
        # Errors during upload raise a DeepLException
        print(error)

def translate_with_glossary(auth_key, target_lang, input_path, output_path, glossary_path):
    translator = deepl.Translator(auth_key)
    with open(glossary_path, "r", encoding="utf-8") as file:
        glossary_id = file.read()
    my_glossary = translator.get_glossary(glossary_id)
    try:
        translator.translate_document_from_filepath(
            input_path,
            output_path,
            source_lang = "EN",
            target_lang = target_lang,
            glossary = my_glossary
        )
    except deepl.DocumentTranslationException as error:
        # If an error occurs during document translation after the document was
        # already uploaded, a DocumentTranslationException is raised. The
        # document_handle property contains the document handle that may be used to
        # later retrieve the document from the server, or contact DeepL support.
        doc_id = error.document_handle.id
        doc_key = error.document_handle.key
        print(f"Error after uploading ${error}, id: ${doc_id} key: ${doc_key}")
    except deepl.DeepLException as error:
        # Errors during upload raise a DeepLException
        print(error)

def translate_with_chinese_glossary(auth_key, target_lang, input_path, output_path, glossary_path):
    translator = deepl.Translator(auth_key)
    try:
        translator.translate_document_from_filepath(
            input_path,
            output_path,
            source_lang= "EN",
            target_lang= target_lang
        )
    except deepl.DocumentTranslationException as error:
        # If an error occurs during document translation after the document was
        # already uploaded, a DocumentTranslationException is raised. The
        # document_handle property contains the document handle that may be used to
        # later retrieve the document from the server, or contact DeepL support.
        doc_id = error.document_handle.id
        doc_key = error.document_handle.key
        print(f"Error after uploading ${error}, id: ${doc_id} key: ${doc_key}")
    except deepl.DeepLException as error:
        # Errors during upload raise a DeepLException
        print(error)
    
    with open(glossary_path, "r", encoding="utf-8") as file:
        glossary = file.read()

    doc = Document(output_path)
    # Loop through all the paragraphs in the document
    for para in doc.paragraphs:
        # Loop through all the runs in the paragraph
        for run in para.runs:
            # Loop through all the entries in the replacements dictionary
            for key, value in glossary.items():
                # Check if the current run contains the current key
                if key in run.text:
                    # Replace the key with the corresponding value
                    run.text = run.text.replace(key, value)

    # Save the updated document
    doc.save(output_path)

def translate_docx(auth_key, input_path, target_lang, output_folder_path, glossary_info_path):
    filename = input_path.split('/')[-1].split('.')[0]
    original_file = f"{output_folder_path}/{filename}.docx"
    translated_files = [original_file]
    for lang in target_lang:
        if lang[0] == "ZH":
            if lang[1] == True:
                output_path = f"{output_folder_path}/{filename} {lang[0]}.docx"
                translated_files.append(output_path)
                if lang[2] == True:
                    glossary_path = f"{glossary_info_path}/glossary_{lang[0]}.txt"
                    translate_with_chinese_glossary(auth_key, lang[0], input_path, output_path, glossary_path)
                elif lang[2] == False:
                    translate_without_glossary(auth_key, lang[0], input_path, output_path)
            elif lang[1] == False:
                pass
            else:
                print("ERROR: True/False is keyed wrongly")
        elif lang[0] == "JA" or lang[0] == "DE":
            if lang[1] == True:
                output_path = f"{output_folder_path}/{filename} {lang[0]}.docx"
                translated_files.append(output_path)
                if lang[2] == True:
                    glossary_path = f"{glossary_info_path}/glossary_id_{lang[0]}.txt"
                    translate_with_glossary(auth_key, lang[0], input_path, output_path, glossary_path)
                elif lang[2] == False:
                    translate_without_glossary(auth_key, lang[0], input_path, output_path)
            elif lang[1] == False:
                pass
            else:
                print("ERROR: True/False is keyed wrongly")
        else:
            print("ERROR: Target language is keyed wrongly")
    shutil.move(input_path, original_file)
    return translated_files