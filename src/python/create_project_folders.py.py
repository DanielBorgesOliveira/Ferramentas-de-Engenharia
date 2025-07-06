import os

# List of folders to create
pairs = [
    {
        "root_folder": r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\1_em_desenvolvimento\BdB211813\Projeto\Pacote 8 - Lavagem de Pecas", 
        "folders": [
            "BdB211813-0000-V-ET0007",
            "BdB211813-0000-V-FD0001",
            "BdB211813-0000-V-FE0011",
        ]
    },
    {
        "root_folder": r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\1_em_desenvolvimento\BdB211813\Projeto\Pacote 9 - HVAC", 
        "folders": [
            "BdB211813-0000-V-ET0010",
            "BdB211813-0000-V-FD0017",
        ]
    }
]

# Loop through each pair and create folders
for pair in pairs:
    root_folder = pair["root_folder"]  # Assign once per pair
    for folder in pair["folders"]:
        folder_path = os.path.join(root_folder, folder)
        try:
            os.makedirs(folder_path, exist_ok=True)
            print(f"✅ Folder created: {folder_path}")
        except Exception as e:
            print(f"❌ Error creating folder {folder_path}: {e}")
