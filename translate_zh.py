from docx import Document
import deepl

auth_key = ""  # Replace with your key
translator = deepl.Translator(auth_key)
source_lang = "EN"
target_lang = "ZH" # DE for German, JA for Japanese and ZH for Chinese (does not work for glossary)

input_path = ".docx" # file to translate
output_path = ".docx" # translated file name

words = ["", "", "", ""] # list of words or phrases to keep
replacements = {}

for word in words:
    result = translator.translate_text(word, target_lang=target_lang)
    replacements[result.text] = f"{result.text} ({word})"

try:
    # Using translate_document_from_filepath() with file paths 
    translator.translate_document_from_filepath(
        input_path,
        output_path,
        source_lang= source_lang,
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

doc = Document(output_path)

# Loop through all the paragraphs in the document
for para in doc.paragraphs:
    # Loop through all the runs in the paragraph
    for run in para.runs:
        # Loop through all the entries in the replacements dictionary
        for key, value in replacements.items():
            # Check if the current run contains the current key
            if key in run.text:
                # Replace the key with the corresponding value
                run.text = run.text.replace(key, value)

# Save the updated document
doc.save(output_path)