import deepl 

auth_key = ""  # Replace with your key
translator = deepl.Translator(auth_key)
source_lang = "EN"
target_lang = "JA" # DE for German, JA for Japanese and ZH for Chinese (does not work for glossary)

input_path = ".docx" # file to translate
output_path = ".docx" # translated file name

glossary_id = "" # Replace with your glossary id
my_glossary = translator.get_glossary(glossary_id)

try:
    # Using translate_document_from_filepath() with file paths 
    translator.translate_document_from_filepath(
        input_path,
        output_path,
        source_lang= source_lang,
        target_lang= target_lang,
        glossary=my_glossary
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