from src.deepl_glossary import create_glossary, create_replacements

auth_key = ""  # Replace with your DeepL API key
word_list = ["This Week's News",   # list of words or phrases to keep
             "Media Releases", 
             "Latest Research"
             ] 
target_lang = "DE" # DE for German, JA for Japanese and ZH for Chinese
glossary_data_path = "./lib/glossary_data"

if target_lang == "ZH":
    create_replacements(auth_key, target_lang, word_list, glossary_data_path)
elif target_lang == "DE" or target_lang == "JA":
    glossary_id = create_glossary(auth_key, target_lang, word_list, glossary_data_path)
else:
    print("ERROR: Target language is keyed wrongly")