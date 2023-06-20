import os

def input_folder(folder_path):
    files = os.listdir(folder_path)
    folder_path = f"{folder_path}/"
    if len(files) == 1:
        file_name = files[0]
        file_path = os.path.join(folder_path, file_name)
        if file_name.endswith(".docx"):
            input_path = file_path
        else:
            input_path = ""
            print("The file in the folder is not in DOCX format.")
    else:
        input_path = ""
        print("The folder does not contain only one file.")
    return input_path