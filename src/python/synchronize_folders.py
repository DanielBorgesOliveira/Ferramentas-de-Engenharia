# https://www.geeksforgeeks.org/autorun-a-python-script-on-windows-startup/
# C:\Users\dboliveira\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\
import os
import shutil
import unicodedata
import logging
import time
import re
import stat

import pdb
DEBUG = False

# Configurable Parameters
FOLDER_PAIRS = [
    (r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\1_em_desenvolvimento\BdB201460\Projeto", 
     r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\6-COMPARTILHADOS\BdB201460\Projeto"),
    
    (r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\1_em_desenvolvimento\BdB211813", 
     r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\6-COMPARTILHADOS\BdB211813"),
]

IGNORED_EXTENSIONS = [".dwl", ".dwl2", ".bak", ".ini", ".log", ".db", ".txt", ".out", ".bakA-001"]  # Extensions to ignore
SYNC_INTERVAL_MINUTES = 10  # Synchronization interval in minutes
LOG_FILE_NAME = "backup_log.txt"

# Logging Setup
def setup_logger(log_file_name, save_to_file = False):
    logger = logging.getLogger("backup_logger")
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(stream_handler)
        
        if save_to_file:
            file_handler = logging.FileHandler(log_file_name, mode='w')
            file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logger.addHandler(file_handler)
    
    logger.info("Logging system initialized.")
    return logger

logger = setup_logger(LOG_FILE_NAME)

# Utility Functions
def normalize_path(path):
    """
    Normalize a path to remove accented characters and support long paths on Windows.
    """
    normalized = unicodedata.normalize('NFC', path)
    
    if os.name == 'nt' and not normalized.startswith('\\\\?\\') and len(normalized) > 260:
        normalized = f"\\\\?\\{normalized}"
    
    return normalized

def force_remove_folder(path):
    def onerror(func, path, exc_info):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception as e:
            logger.error(f"Erro forçando remoção de: {path}. Detalhes: {e}")
            log_folder_diagnostics(path)
    shutil.rmtree(path, onerror=onerror)

def log_folder_diagnostics(path):
    logger.info(f"Diagnóstico de pasta antes da exclusão: {path}")
    try:
        stats = os.stat(path)
        logger.info(f"Atributos: {stats.st_mode} | Somente leitura: {not os.access(path, os.W_OK)}")
    except Exception as e:
        logger.error(f"Erro ao acessar atributos da pasta: {e}")

# Synchronization Function
def synchronize_folders(source_folder, destination_folder):
    if DEBUG: pdb.set_trace()
    if not os.path.exists(source_folder):
        logger.error(f"Source folder '{source_folder}' does not exist.")
        return
    
    logger.info(f"Synchronizing from '{source_folder}' to '{destination_folder}'")
    
    dest_files_to_check = set()
    
    # Copia arquivos que existem na origem mas não no destino
    for root, dirs, files in os.walk(source_folder):
        # Pega o caminho relativo da pasta raiz em relação a pasta source_folder
        relative_path = os.path.relpath(root, source_folder)
        
        # Concatena o caminho para formar o caminho da pasta destino
        if relative_path == ".":
            dest_path = destination_folder
        else:
            dest_path = os.path.join(destination_folder, relative_path)
        
        # Cria a pasta no caminho, caso ela não exista.
        os.makedirs(dest_path, exist_ok=True)
        
        for file in files:
            # Pega o caminho absoluto do arquivo raiz
            source_file = normalize_path(os.path.join(root, file))
            
            # Pega o caminho absoluto do arquivo de destino
            dest_file = normalize_path(os.path.join(dest_path, file))
            file_extension = os.path.splitext(file)[1].lower()
            
            # Verifica se o arquivo deve ser ignorado.
            pattern = re.compile(rf"{re.escape(file_extension)}")
            if list(filter(pattern.search, IGNORED_EXTENSIONS)):
                #logger.info(f"O arquivo foi ignorado. Arquivo {source_file}. Extensão: {file_extension}")
                continue
            
            dest_files_to_check.add(dest_file)
            
            try:
                if not os.path.exists(dest_file) or os.path.getmtime(source_file) > os.path.getmtime(dest_file):
                    shutil.copy2(source_file, dest_file)
                    logger.info(f"Copied: {os.path.join(relative_path, file)}", exc_info=True)
            except Exception as e:
                logger.error(f"Error copying {source_file}: {e}")
    
    # Deleta arquivos que existem no destino mas não existem na origem
    for root, dirs, files in os.walk(destination_folder):
        # Pega o caminho relativo da pasta destino em relação a pasta source_folder
        relative_path = os.path.relpath(root, destination_folder)
        
        # Concatena o caminho para formar o caminho da pasta destino
        if relative_path == ".":
            dest_path = destination_folder
        else:
            dest_path = os.path.join(destination_folder, relative_path)
        
        for file in files:
            dest_file = normalize_path(os.path.join(dest_path, file))
            if dest_file.lower() not in (f.lower() for f in dest_files_to_check):
                try:
                    os.remove(dest_file)
                    logger.info(f"Deleted: {os.path.join(relative_path, file)}")
                except Exception as e:
                    logger.error(f"Error deleting file: {dest_file}. Error: {e}", exc_info=True)
    
    # Deleta pastas que existem no destino mas não existem na origem
    for root, dirs, files in os.walk(destination_folder, topdown=False):
        # Pega o caminho relativo da pasta destino em relação a pasta source_folder
        relative_path = os.path.relpath(root, destination_folder)
        
        # Concatena o caminho para formar o caminho da pasta na origem
        if relative_path == ".":
            source_equiv = source_folder
        else:
            source_equiv = os.path.join(source_folder, relative_path)
        
        # Se a pasta não exitir na origem a pasta é deletada no destino.
        if not os.path.exists(source_equiv):
            try:
                force_remove_folder(root)
                logger.info(f"Deleted folder: {relative_path}")
            except Exception as e:
                logger.error(f"Error deleting folder: {root}. Error: {e}", exc_info=True)

# Main Loop
if __name__ == "__main__":
    while True:
        for source, destination in FOLDER_PAIRS:
            try:
                synchronize_folders(source, destination)
            except Exception as e:
                logger.error(f"Error during synchronization of {source} to {destination}: {e}")
        
        logger.info(f"Synchronization cycle complete. Waiting {SYNC_INTERVAL_MINUTES} minutes for the next run.")
        time.sleep(SYNC_INTERVAL_MINUTES * 60)
