import deepl, shutil, pickle, os
from docx import Document

config = {
    "auth_key" : "daa69283-7136-c9a5-c283-38cbd645e38e", # Copy key from DeepL
    "input_folder" : "./input_folder",
    "output_folder" : "./docx_folder",
    "languages" : {
        "DE": {"translate": True, "glossary": False},
        "JA": {"translate": False, "glossary": False},
        "ZH": {"translate": False, "glossary": False}
    },
    "glossary_folder" : "./lib/glossary_data"
}

def check_config():
    assert type(config["auth_key"]) == str, "Auth_key is not string"
    for lang, operation in config["languages"].items():
        assert lang in ["DE", "JA", "ZH"], "Wrong language code for German, Japanese and Chinese"
        assert type(operation["translate"]) == bool, "Translation flag is not boolean"
        assert type(operation["glossary"]) == bool, "Glossary flag is not boolean"

def handle_error(error):
    # Handle errors associated with DeepL API

    if isinstance(error, deepl.DocumentTranslationException):
        '''If an error occurs during document translation after the document was
        already uploaded, a DocumentTranslationException is raised. The
        document_handle property contains the document handle that may be used to
        later retrieve the document from the server, or contact DeepL support.'''

        doc_id = error.document_handle.id
        doc_key = error.document_handle.key
        print(f"{error}\nError after uploading {error}, id: {doc_id} key: {doc_key}")
    elif isinstance(error, deepl.DeepLException):
        # Errors during upload raise a DeepLException
        print(f"{error}\nErrors during upload")
    else:
        print(error)

def extract_docxfile(input_folder):
    # Extract DOCX file from input folder

    files = os.listdir(input_folder)
    if len(files) == 1:
        file_name = files[0]
        file_path = os.path.join(input_folder, file_name)
        if file_name.endswith(".docx"):
            return file_path
        else:
            print("The file in the folder is not in DOCX format.")
            exit()
    else:
        print("The folder does not contain exactly one file.")
        exit()

def extract_filename(file_path):
    '''Extract file name'''

    file_name = file_path.split('/')[-1]
    file_name = file_name.split('\\')[-1].split('.')[0]
    return file_name


def translate_docx(auth_key, file_path, languages, output_folder, glossary_folder):
    '''Translate DOCX'''

    file_name = extract_filename(file_path)

    translated_files_list = []

    translator = deepl.Translator(auth_key)

    for lang, operation in languages.items():
        try: 
            if lang == "DE" or lang == "JA":
                if operation["translate"] == True:
                    output_file = f"{output_folder}/{file_name} {lang}.docx"
                    translated_files_list.append(output_file)
                    
                    if operation["glossary"] == True:
                        glossary_path = f"{glossary_folder}/glossary_id_{lang}.txt"
                        
                        # Read glossary ID
                        try:
                            with open(glossary_path, "r", encoding="utf-8") as file:
                                glossary_id = file.read()
                                my_glossary = translator.get_glossary(glossary_id)

                            # Translate to DE or JA with glossary
                            try:
                                translator.translate_document_from_filepath(
                                    file_path,
                                    output_file,
                                    source_lang = "EN",
                                    target_lang = lang,
                                    glossary = my_glossary
                                )
                            except Exception as error:
                                handle_error(error)
                        except Exception as error:
                            print(f"Error: Could not load glossary for {lang} from {glossary_path}.")
                            pass

                    elif operation["glossary"] == False:
                        
                        # Translate to DE or JA without glossary
                        try:
                            translator.translate_document_from_filepath(
                                file_path,
                                output_file,
                                source_lang = "EN",
                                target_lang = lang
                            )
                        except Exception as error:
                            handle_error(error)
                        
                elif operation["translate"] == False:
                    # No need to translate for this language
                    pass
                
            elif lang == "ZH":
                if operation["translate"] == True:
                    output_file = f"{output_folder}/{file_name} {lang}.docx"
                    translated_files_list.append(output_file)

                    # Translate to Chinese
                    try:
                        translator.translate_document_from_filepath(
                            file_path,
                            output_file,
                            source_lang = "EN",
                            target_lang = lang
                        )
                    except Exception as error:
                        handle_error(error)

                    if operation["glossary"] == True:
                        glossary_path = f"{glossary_folder}/glossary_{lang}.pkl" 
                        
                        # Dictionary of replacements to be executed
                        try:
                            with open(glossary_path, 'rb') as f:
                                glossary = pickle.load(f)

                            # Execute replacements
                            with open(output_file, 'rb') as f:
                                doc = Document(f)

                            for para in doc.paragraphs:
                                for run in para.runs:
                                    for key, value in glossary.items():
                                        if key in run.text:
                                            run.text = run.text.replace(key, value)

                            with open(output_file, 'wb') as f:
                                doc.save(f)
                        except (FileNotFoundError, pickle.UnpicklingError):
                            print(f"Error: Could not load glossary for {lang} from {glossary_path}.")
                            pass
                        
                    elif operation["glossary"] == False:
                        # No need to do replacements
                        pass
                    
                        pass
                elif operation["translate"] == False:
                    # No need to translate for this language
                    pass
                
        except Exception as error:
            print(error)

    # Move original file to output folder, thus clear input folder
    moved_original_file = f"{output_folder}/{file_name}.docx"
    shutil.move(file_path, moved_original_file)
    translated_files_list.append(moved_original_file)

    return translated_files_list

if __name__ == "__main__":
    check_config()
    file_path = extract_docxfile(config["input_folder"])
    translate_docx(config["auth_key"], file_path, config["languages"], config["output_folder"], config["glossary_folder"])
    print("Translated file.")