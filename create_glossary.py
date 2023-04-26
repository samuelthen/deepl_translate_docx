import requests
import deepl

auth_key = ""  # Replace with your DeepL API key
translator = deepl.Translator(auth_key)
source_lang = "EN"
target_lang = "JA" # DE for German, JA for Japanese and ZH for Chinese (does not work for glossary)

# Set a glossary
words = ["", "", "", ""] # list of words or phrases to keep
entries = {}
for word in words:
    result = translator.translate_text(word, target_lang=target_lang)
    entries[word] = f"{result.text} ({word})"

entries_tsv = "\n".join([f"{key}\t{value}" for key,value in entries.items()])
create_glossary_url = "https://api-free.deepl.com/v2/glossaries"
headers = {"Authorization": f"DeepL-Auth-Key {auth_key}"}
create_glossary_data = {
    "name": "My Glossary",
    "source_lang": source_lang,
    "target_lang": target_lang,
    "entries": entries_tsv,
    "entries_format": "tsv",
}
response = requests.post(create_glossary_url,data=create_glossary_data,headers=headers,)

glossary_id = response.json()["glossary_id"]
print(glossary_id)