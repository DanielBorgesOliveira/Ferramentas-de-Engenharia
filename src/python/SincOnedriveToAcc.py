#!/usr/bin/python

#"C:\Users\dboliveira\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\0-GERENCIAMENTO\Script\SincronizeOnedriveToAcc.py"

import shutil
import psutil
import os
from datetime import datetime
import re
import json
import time
import threading
import tempfile
from pathlib import Path
import logging
import logging.handlers
import queue
import stat

onedrive_base_folder = Path(r"C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos")
pattern = re.compile(r'(\D\D\D\d\d\d\d\d\d)', flags = re.IGNORECASE)
cpu_usage = 50 # [%]
IGNORED_EXTENSIONS = {".dwl", ".dwl2", ".bak", ".ini", ".log", ".db", ".out", ".bakA-001"}  # Extensions to ignore
log_file_name = onedrive_base_folder / "backup_log.txt"
projects_to_ignore = [
    'BdB211804', # Projeto finalizado
]
flows = [
    "1_em_desenvolvimento",
    "2_em_verificacao",
    "3_em_atendimento_de_comentarios",
    "4_em_emissao",
    "5_concluido",
]

# Mapeamento dos diretórios de destino
acc_base_folder = {
    'BdB240102': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240102-Anglo-MisturDeLama\Project Files\Geral\V"),
    'BdB230203': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\230203-Anglo-DFsFiltrag\Project Files\Geral\V"),
    'BdB211806': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\SAMARCO\Project Files\BdB2118\BdB211806_SAMARCO_TRAS_REJ_C2\V"),
    'BdB210112': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\210112-Anglo-LamodutoCMD\Project Files\Geral\V"),
    'BdB240300': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240300-Anglo-PfsaPostoCMD\Project Files\Geral\V"),
    'BdB210113': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\210113-Anglo-ReportBomMine\Project Files\Geral\V"),
    'BdB201431': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201431-Vale-AdeqRejeitTBO\Project Files\Geral\V"),
    'BdB201452': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201452-Vale-BombRejeArea8\Project Files\Geral\V"),
    'BdB220500': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\220500-CBMM-EDR9\Project Files\Geral\V"),
    'BdB200301': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\220500-CBMM-EDR9\Project Files\Geral\V"),
    'BdB200305': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\220500-CBMM-EDR9\Project Files\Geral\V"),
    'BdB200302': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\C B M M\Project Files\BdB2003\BdB200302_CBMM_EXTRAÇÃO_MEC\V"),
    'BdB201459': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201459-Vale-AdeqFerejViga\Project Files\Geral\V"),
    'BdB211810': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211810-Samarco-RevampC1\Project Files\Geral\V"),
    'BdB201454': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201454-Vale-NovRotRejVig\Project Files\Geral\V"),
    'BdB201451': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201451-Vale-RedFeBrucutu\Project Files\Geral\V"),
    'BdB211622': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211622-CBMM-EtelIII-EDR9\Project Files\Geral\V"),
    'BdB201460': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201460-Vale-Adeq_PNR_SPCI-TMPM\Project Files\Geral\V"),
    "BdB211811": Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211810-Samarco-RevampC1\Project Files\Geral\V"),
    'BdB240101': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240101-Anglo-SistBombCava\Project Files\Geral\V"),
    'BdB210109': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\ANGLOAMERICAN\Project Files\2101\BdB210109_ANGLO_PLAN_FILTR_REJ\V"),
    'BdB211813': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211813-Samarco-InstPDER\Project Files\Geral\V"),
    'BdB240205': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240205-Anglo-MatrizBloq\Project Files\Geral\V"),
    'BdB240104': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240104-Anglo-PFSABombLama\Project Files\Geral\V"),
    'BdB201457': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201457-Vale-ATORejVGR1\Project Files\Geral\V"),
    'BdB240105': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240105-Anglo-TradeOffBombCava\Project Files\Geral\V"),
    'BdB211615': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211615-CBMM-UnidComplementEDR\Project Files\Geral\V"),
    'BdB211624': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211624-CBMM-EDR9-ServAdic\Project Files\Geral\V"),
    'BdB211815': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211624-CBMM-EDR9-ServAdic\Project Files\Geral\V"),
    'BdB211805': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\211624-CBMM-EDR9-ServAdic\Project Files\Geral\V"),
    'BdB240109': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\240109-Anglo-MistLamasBas\Project Files\01_WIP\Geral\V"),
    'BdB201466': Path(r"C:\Users\dboliveira\DC\ACCDocs\BRASS Engineering\201466-Vale-OtimCircuVGR1\Project Files\Geral\V"),
}

r"""
acc_base_folder = {
    'BdB201454': Path(r"C:\v\BdB201454"),
    'BdB210109': Path(r"C:\v\BdB210109"),
    'BdB210112': Path(r"C:\v\BdB210112"),
    'BdB200305': Path(r"C:\v\BdB200305"),
    'BdB201451': Path(r"C:\v\BdB201451"),
    'BdB240300': Path(r"C:\v\BdB240300"),
    'BdB211804': Path(r"C:\v\BdB211804"),
    'BdB201452': Path(r"C:\v\BdB201452"),
    'BdB210113': Path(r"C:\v\BdB210113"),
    'BdB201431': Path(r"C:\v\BdB201431"),
    'BdB200302': Path(r"C:\v\BdB200302"),
    'BdB200301': Path(r"C:\v\BdB200301"),
    'BdB240101': Path(r"C:\v\BdB240101"),
    'BdB211806': Path(r"C:\v\BdB211806"),
    'BdB211622': Path(r"C:\v\BdB211622"),
    'BdB201459': Path(r"C:\v\BdB201459"),
    'BdB220500': Path(r"C:\v\BdB220500"),
    'BdB201460': Path(r"C:\v\BdB201460"),
    'BdB240102': Path(r"C:\v\BdB240102"),
    'BdB211810': Path(r"C:\v\BdB211810"),
    'BdB211811': Path(r"C:\v\BdB211811"),
    'BdB230203': Path(r"C:\v\BdB230203"),
    'BdB211813': Path(r"C:\v\BdB211813"),
    'BdB240205': Path(r"C:\v\BdB240205"),
    'BdB240104': Path(r"C:\v\BdB240104"),
    'BdB201457': Path(r"C:\v\BdB201457"),
    'BdB240105': Path(r"C:\v\BdB240105"),
    'BdB211615': Path(r"C:\v\BdB211615"),
}
"""

def setup_logger(log_file_name):
    # Clear the log file at the start of the script
    with open(log_file_name, 'w'):
        pass
    
    # Criar uma fila para armazenar logs
    log_queue = queue.Queue()
    
    # Criar o logger principal
    logger = logging.getLogger("backup_logger")
    logger.setLevel(logging.INFO)
    
    # Criar o manipulador de arquivo
    file_handler = logging.FileHandler(log_file_name, mode='w')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    
    # Criar um manipulador de fila
    queue_handler = logging.handlers.QueueHandler(log_queue)
    logger.addHandler(queue_handler)
    
    # Criar o listener da fila que escreverá os logs no arquivo
    queue_listener = logging.handlers.QueueListener(log_queue, file_handler)
    queue_listener.start()  # Iniciar a thread do listener
    
    # Criar um manipulador de console para visualizar logs na tela
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(stream_handler)
    
    logger.info("Logging system initialized.")
    return logger, queue_listener

# Configurar logger
logger, log_listener = setup_logger(log_file_name)

def listdirs(folder, pattern):
    output = []
    for it in os.scandir(folder):
        if it.is_dir():
            pattern_result = pattern.findall(it.path)
            if pattern_result and pattern_result[0] not in output:
                output.append(pattern_result[0])
    return(output)

def safe_remove_file(file_path, retries=3, delay=1):
    last_error = None

    for attempt in range(retries):
        try:
            if os.path.exists(file_path):
                # Remove somente leitura, se necessário
                os.chmod(file_path, stat.S_IWRITE)
                os.remove(file_path)
            return True
        except Exception as e:
            last_error = e
            time.sleep(delay)

    if os.path.exists(file_path):
        logger.warning(f"Failed to delete file after {retries} attempts: {file_path}, last error: {last_error}")
        return False
    else:
        return True  # Foi removido com sucesso por outro processo durante os retries

def safe_remove_dir(path, retries=3, delay=1):
    def on_rm_error(func, path, exc_info):
        # Remove somente leitura e tenta novamente
        os.chmod(path, stat.S_IWRITE)
        try:
            func(path)
        except Exception as e:
            logger.error(f"Failed to forcibly delete {path}: {e}")

    last_error = None

    for attempt in range(retries):
        try:
            if os.path.exists(path):
                shutil.rmtree(path, onerror=on_rm_error)
            return True
        except Exception as e:
            last_error = e
            time.sleep(delay)
    
    logger.warning(f"Failed to delete directory after {retries} attempts: {path}, last error: {last_error}")
    return False

def backup_folder(sync_item):
    """
    Synchronizes destination_folder with source_folder.
    Copies newer files from source_folder to destination_folder,
    and deletes files/folders in destination_folder that do not exist in source_folder.
    
    Parameters:
    - source_folder (str): Path of the folder to back up.
    - destination_folder (str): Path of the folder where the backup will be stored.
    """
    
    source_folder = sync_item["From"]
    destination_folder = sync_item["To"]
    
    # Ensure source exists and is a directory
    if not os.path.isdir(source_folder):
        logger.info(f"Folder does not exist or is not a directory: '{source_folder}'")
        return
    
    # Ensure destination folder exists, create it if not.
    os.makedirs(destination_folder, exist_ok=True)
    
    # Delete files and folders in destination that are not in source
    for dirpath, dirnames, filenames in os.walk(destination_folder, topdown=False):
        # Check if each directory in the destination exists in the source
        rel_path = os.path.relpath(dirpath, destination_folder)
        src_dirpath = os.path.join(source_folder, rel_path)
        
        # Delete directories not present in source
        if not os.path.exists(src_dirpath):
            try:
                safe_remove_dir(dirpath)
                #logger.info(f"Deleted directory: {dirpath}")
            except Exception as error:
                logger.error(f"Error deleting directory: dirpath: {dirpath}, error {error}")
            continue
        
        # Delete files not present in source
        for filename in filenames:
            dest_file = os.path.join(dirpath, filename)
            src_file = os.path.join(src_dirpath, filename)
            if not os.path.exists(src_file):
                try:
                    safe_remove_file(dest_file)
                    #logger.info(f"Deleted file: {dest_file}")
                except Exception as error:
                    logger.error(f"Error deleting file: dest_file: {dest_file}, error {error}")
    
    # Copy new and updated files from source to destination
    for dirpath, dirnames, filenames in os.walk(source_folder):
        # Create corresponding directories in the destination
        rel_path = os.path.relpath(dirpath, source_folder)
        dest_dir = os.path.join(destination_folder, rel_path)
        os.makedirs(dest_dir, exist_ok=True)
        
        for filename in filenames:
            _, ext = os.path.splitext(filename)
            if ext.lower() in IGNORED_EXTENSIONS:
                #logger.info(f"Ignored file due to extension: {filename}")
                continue
            
            src_file = os.path.join(dirpath, filename)
            dest_file = os.path.join(dest_dir, filename)
            
            # Log para verificar os caminhos antes da cópia
            if not os.path.exists(src_file):
                logger.info(f"File not found (skipping): {src_file}")
            elif not os.path.isdir(dest_dir):
                logger.info(f"Folder in the destination was not created (skipping): {dest_dir}")
            else:
                # Continue com a cópia se os caminhos forem válidos
                if os.path.exists(dest_file):
                    src_mtime = os.path.getmtime(src_file)
                    dest_mtime = os.path.getmtime(dest_file)
                    if src_mtime > dest_mtime:
                        try:
                            shutil.copy2(src_file, dest_file)
                            logger.info(f"Updated file: {dest_file}")
                        except Exception as error:
                            logger.error(f"Error updating file src_file: {src_file}, dest_file: {dest_file}, error: {error}")
                    # DEBUG
                    #else:
                    #    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{datetime.now().microsecond // 1000:03d}- INFO - Skiped file: src_file: {src_file}, dest_file: {dest_file}")
                else:
                    try:
                        # DEBUG
                        #input(f"DEBUG: src_file: {src_file}, dest_file: {dest_file}")
                        shutil.copy2(src_file, dest_file)
                        logger.info(f"Copied new file: {dest_file}")
                    except Exception as error:
                        logger.error(f"Error copying file: src_file: {src_file}, dest_file: {dest_file}, error {error}")

def throttled_backup(sync_item):
    """Backs up a folder while dynamically throttling based on CPU usage."""
    backup_folder(sync_item)
    
    # Throttle if CPU usage exceeds the desired threshold
    while psutil.cpu_percent(interval=0.1) > cpu_usage:
        time.sleep(0.5)  # Brief pause to reduce CPU load

def main():
    projects = []
    for flow in flows:
        projects.append(listdirs(folder = fr'{onedrive_base_folder}\{flow}', pattern = pattern))

    projects = list(set([x for y in projects for x in y]))

    projects = [project for project in projects if project not in projects_to_ignore]

    # Create a temp dir
    temp_dir = tempfile.TemporaryDirectory()

    sync_map = []
    for flow in flows:
        projects_in_flow = listdirs(folder = fr'{onedrive_base_folder}\{flow}', pattern = pattern)
        for project in projects:
            if project in projects_in_flow:
                sync_map.append({
                    "From": fr'{onedrive_base_folder}\{flow}\{project}',
                    "To": fr'{acc_base_folder[project]}\FolderSync\{flow}'
                    #"To": fr'{temp_dir.name}\{project}\{flow}' # DEBUG
                })

    # Using threading to handle parallel backups
    threads = []
    for sync in sync_map:
        # Execução em serie
        #backup_folder(sync)
        
        # Execução em paralelo com controle de nível de processamento
        #thread = threading.Thread(target=throttled_backup, args=(sync,))
        
        # Execução em paralelo sem controle de nível de processamento
        thread = threading.Thread(target=backup_folder, args=(sync,))
        
        threads.append(thread)
        thread.start()

    # Wait for all threads to complete
    for thread in threads:
        thread.join()

    # No final do script, antes de sair, parar o listener para evitar threads em execução:
    log_listener.stop()

#print(json.dumps(sync_map, indent=4))
#backup_folder(sync_map[0]["From"], sync_map[0]["To"])

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"\n[ERROR] Ocorreu uma exceção: {e}")

    input("Pressione Enter para sair...")

