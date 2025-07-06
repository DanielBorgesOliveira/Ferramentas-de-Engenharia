import zipfile
import os
from oletools.olevba import VBA_Parser, FileOpenError

def extract_vba_from_dotm(dotm_path, export_dir):
    # Ensure the export directory exists
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)
    
    # Step 1: Extract the .dotm file as a ZIP archive
    with zipfile.ZipFile(dotm_path, 'r') as zip_ref:
        zip_ref.extractall(export_dir)
    
    # Step 2: Locate the vbaProject.bin file
    vba_project_path = os.path.join(export_dir, 'word', 'vbaProject.bin')
    if not os.path.exists(vba_project_path):
        print("vbaProject.bin not found in the .dotm file.")
        return
    
    # Step 3: Use oletools to parse the vbaProject.bin file
    try:
        vba_parser = VBA_Parser(vba_project_path)
        if vba_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                if vba_code is not None:
                    export_path = os.path.join(export_dir, vba_filename)
                    with open(export_path, 'w', encoding='latin1') as f:
                        f.write(vba_code)
                        print(f"Exported: {vba_filename} to {export_path}")
        else:
            print("No VBA macros found in the vbaProject.bin file.")
    except FileOpenError as e:
        print(f"Error opening vbaProject.bin: {e}")
    
    # Clean up extracted files
    print("Extraction completed.")

# Define the path to the .dotm file and the export directory
dotm_path = r"C:\Users\daniel\Desktop\1\FinishHim - v0.5.dotm"
export_dir = r"C:\Users\daniel\Desktop\1"

# Run the extract function
extract_vba_from_dotm(dotm_path, export_dir)