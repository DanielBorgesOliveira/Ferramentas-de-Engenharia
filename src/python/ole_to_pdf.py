import os

str1 = b'%PDF-'  # Begin PDF
str2 = b'%%EOF'  # End PDF

BasePath = r"C:\Users\dboliveira\Downloads"

def find_files(extension, search_path):
    result = []
    for root, dirs, files in os.walk(search_path):
        for file in files:
            if file.endswith(extension):
                 result.append(
                    {
                        'InputFile': file,
                        'InputDir': root,
                        'OutputFile': file.split(".")[0]+"."+"pdf",
                        'OutputDir': root,
                    }
                )
    return result

Files = find_files(".bin", BasePath)

for file in Files:
    with open(file['InputDir']+"\\"+file['InputFile'], 'rb') as f:
        binary_data = f.read()
    
    # Convert BYTE to BYTEARRAY
    binary_byte_array = bytearray(binary_data)
    
    # Find where PDF begins
    result1 = binary_byte_array.find(str1)
    
    # Remove all characters before PDF begins
    del binary_byte_array[:result1]
    
    # Find where PDF ends
    result2 = binary_byte_array.find(str2)
    
    # Subtract the length of the array from the position of where PDF ends (add 5 for %%OEF characters)
    # and delete that many characters from end of array
    to_remove = len(binary_byte_array) - (result2 + 5)
    
    del binary_byte_array[-to_remove:]
    
    with open(file['OutputDir']+"\\"+file['OutputFile'], 'wb') as fout:
        fout.write(binary_byte_array)
