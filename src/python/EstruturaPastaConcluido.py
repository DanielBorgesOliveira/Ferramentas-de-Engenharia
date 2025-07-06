import os
import shutil

Dates = [
    "15-07-2024",
    "16-04-2024",
    "17-05-2024",
    "17-06-2024",
    "18-06-2024",
    "18-07-2024",
    "18-09-2024",
    "19-07-2024",
    "19-08-2024",
    "21-04-2024",
    "21-05-2024",
    "22-05-2024",
    "23-09-2024",
    "24-07-2024",
    "24-09-2024",
    "26-03-2024",
    "26-04-2024",
    "27-08-2024",
    "29-07-2024",
]

for Date in Dates:
    # Define the source and destination folders
    source_folder = r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\5 - Concluido\\"+Date
    destination_folder = r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\5 - Concluido"
    
    # Ensure the destination folder exists
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    #From: 5 - Concluido\06-06-2024\BdB201459-0000-V-RL0001.pdf. To: 5 - Concluido\BdB201459\BdB201459-0000-V-RL0001\06-06-2024\BdB201459-0000-V-RL0001.pdf
    # List all files in the source folder
    for filename in os.listdir(source_folder):
        file_path = os.path.join(source_folder, filename)
        # Ensure it's a file, not a directory
        if os.path.isfile(file_path):
            # Get the file name without extension
            file_name_without_ext = os.path.splitext(filename)[0]
            # Create a new folder in the destination folder named after the file
            new_folder = os.path.join(destination_folder, file_name_without_ext, Date)
            if not os.path.exists(new_folder):
                os.makedirs(new_folder)
                print(f"Created folder: {new_folder}")
            # Move the file to the new folder
            destination_file_path = os.path.join(new_folder, filename)
            shutil.move(file_path, destination_file_path)
            print(f"Moved file: {filename} to {new_folder}")
    
    #From: 5 - Concluido\06-06-2024\BdB201459-0000-V-RL0001\BdB201459-0000-V-RL0001.pdf. To: 5 - Concluido\BdB201459\BdB201459-0000-V-RL0001\06-06-2024\BdB201459-0000-V-RL0001.pdf
    for dir_name in os.listdir(source_folder):
        from_path = os.path.join(source_folder, dir_name)
        if os.path.isdir(from_path):
            # Original folder with our documents
            document_folder = os.path.basename(os.path.normpath(from_path))
            
            # Project name
            project = document_folder.split("-")[0]
            
            # Destiantion folder to put our documents.
            to_path = os.path.join(destination_folder, project, document_folder, Date)
            
            # Ensure the destination to_path exists
            if not os.path.exists(to_path):
                os.makedirs(to_path)
            
            for file_name in os.listdir(from_path):
                shutil.move(os.path.join(from_path, file_name), to_path)
            
            os.rmdir(from_path)
            
            #input(f'From Folder: {from_path}\nTo Folder: {to_path}')





