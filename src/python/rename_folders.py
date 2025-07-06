import os
import re

def rename_folders(root_dir):
    # This pattern matches directories with the format "DD-MM-YYYY"
    date_pattern = re.compile(r'^(\d{2})-(\d{2})-(\d{4})$')
    
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for dirname in dirnames:
            # Check if the directory name matches the date pattern
            if date_pattern.match(dirname):
                new_name = '-'.join(dirname.split('-')[::-1])  # Reformat to "YYYY-MM-DD"
                old_path = os.path.join(dirpath, dirname)
                new_path = os.path.join(dirpath, new_name)
                os.rename(old_path, new_path)
                print(f'Renamed "{old_path}" to "{new_path}"')

# Example usage:
# Specify the root directory to start from
root_directory = r'C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\5_concluido'
rename_folders(root_directory)
