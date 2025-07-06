import os
import shutil
import re
import json
import time

folders = [
    {
        "bucket": "2 - Em Verificacao",
        "base_folder": r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\2 - Em Verificacao",
        "files": [
            "file": "",
            "folder": "",
            "modification_date": "",
        ],
    },
    {
        "bucket": "3 - Em Atendimento de Comentarios",
        "base_folder": r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\3 - Em Atendimento de Comentarios",
        "files": [
            "file": "",
            "folder": "",
            "modification_date": "",
        ],
    },
    {
        "bucket": "4 - Em Emissao",
        "base_folder": r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\4 - Em Emissao",
        "files": [
            "file": "",
            "folder": "",
            "modification_date": "",
        ],
    },
    {
        "bucket": "5 - Concluido",
        "base_folder": r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\5 - Concluido",
        "files": [
            "file": "",
            "folder": "",
            "modification_date": "",
        ],
    },
]

pattern = re.compile(r'(\D\D\D\d\d\d\d\d\d-\d\d\d\d-\D-\D\D\d\d\d\d)', flags = re.IGNORECASE)

def list_documents(folder, pattern):
    global temp
    for it in os.scandir(folder):
        if it.is_dir():
            pattern_result = pattern.findall(it.path)
            if pattern_result and pattern_result[0] not in temp:
                temp.append({
                    "file": pattern_result[0],
                    "folder": it.path,
                    "modification_date": time.strftime('%Y-%m-%d', time.gmtime(os.path.getmtime(it.path))),
                })
            list_documents(it, pattern)
    return(temp)

temp = []
for i in range(len(folders)):
    folders[i]["files"] = list_documents(folder = folders[i]["base_folder"], pattern = pattern)
    temp = []

# ***************** Delete Files in "5 - Concluido" if in another bucket ***************** #
files_concluido = [x["file"] for x in folders[-1]["files"]]
indexes_to_remove = []
for i in range(len(folders) - 1):  # Iterate through all buckets except "5 - Concluido"
    for file in folders[i]["files"]:
        # Check if the file is in files_concluido
        if file["file"] in files_concluido:
            print(f"File '{file['file']}' found in both bucket {i} and '5 - Concluido'")
            index_to_remove = next((idx for idx, f in enumerate(folders[-1]["files"]) if f["file"] == file["file"]), None)
            if index_to_remove is not None:
                print(f"Removing file '{file['file']}' from '5 - Concluido' at index {index_to_remove}")
                indexes_to_remove.append(index_to_remove)

for index in sorted(indexes_to_remove, reverse=True):
    del folders[-1]["files"][index]
# ***************** ****** ***** ** ** * ********** ** ** ******* ****** ***************** #

#Files = []
#for i in range(len(folders)):
#    for j in range(len(folders[i]["files"])):
#        Files.append({
#            "File": folders[i]["files"][j],
#            "bucket": folders[i]["bucket"]
#        })

#print(json.dumps(Files, indent=4))

for folder in folders:
    for file in folder["files"]:
        print(f'{file["file"]},{file["modification_date"]},{file["modification_date"]},{file["modification_date"]},{folder["bucket"]}')







