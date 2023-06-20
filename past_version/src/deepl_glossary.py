import requests
import deepl
import json

def create_glossary(auth_key, target_lang, word_list, glossary_data_path):    # for German and Japanese
    translator = deepl.Translator(auth_key)
    entries = {}
    for word in word_list:
        result = translator.translate_text(word, target_lang=target_lang)
        entries[word] = f"{result.text} ({word})"

    entries_tsv = "\n".join([f"{key}\t{value}" for key,value in entries.items()])
    create_glossary_url = "https://api-free.deepl.com/v2/glossaries"
    headers = {"Authorization": f"DeepL-Auth-Key {auth_key}"}
    create_glossary_data = {
        "name": "My Glossary",
        "source_lang": "EN",
        "target_lang": target_lang,
        "entries": entries_tsv,
        "entries_format": "tsv",
    }
    response = requests.post(create_glossary_url,data=create_glossary_data,headers=headers)

    glossary_id = response.json()["glossary_id"]
    print(f"Glossary id: {glossary_id}")
    
    with open(f"{glossary_data_path}/glossary_id_{target_lang}.txt", "w") as file:
        json.dump(glossary_id, file)
    with open(f"{glossary_data_path}/glossary_{target_lang}.txt", "w") as file:
        json.dump(entries, file)
    
    return glossary_id

def create_replacements(auth_key, target_lang, word_list, glossary_data_path):
    entries = {}
    translator = deepl.Translator(auth_key)
    for word in word_list:
        result = translator.translate_text(word, target_lang=target_lang)
        entries[result.text] = f"{result.text} ({word})"
    with open(f"{glossary_data_path}/glossary_{target_lang}.txt", "w") as file:
        json.dump(entries, file)

