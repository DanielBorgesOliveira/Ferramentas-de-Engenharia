import win32com.client
import re
import os

def find_files(folder, pattern):
    result = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            pattern_result = pattern.findall(file)
            #if file.endswith(pattern):
            if pattern_result:
                 result.append(
                    {
                        'InputDir': root,
                        'InputFile': file,
                    }
                )
    return result

pattern = re.compile(r'(.\.doc?(.))', flags = re.IGNORECASE)
Files = find_files(
    pattern = pattern, 
    folder = r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\Desktop\1\BdB201452-0000-V-MC0001"
)

word = win32com.client.Dispatch("Word.Application")
for file in Files:
    doc = word.Documents.Open(os.path.join(file['InputDir'], file['InputFile']))
    editing_time = doc.BuiltInDocumentProperties("Total Editing Time")
    print(f"File: {file['InputFile']}. Total Editing Time: {editing_time} minutes")
    doc.Close(False)

word.Quit()
