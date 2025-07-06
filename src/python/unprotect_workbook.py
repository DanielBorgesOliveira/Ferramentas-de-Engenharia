#pip install msoffcrypto-tool openpyxl

import msoffcrypto
import openpyxl
import zipfile
import shutil
import os
import re

def unprotect_workbook(excel_file):
    # Step 1: Rename the original .xlsx file to .zip
    zip_file = excel_file.replace(".xlsx", ".zip")
    zip_file = excel_file.replace(".xlsm", ".zip")
    os.rename(excel_file, zip_file)
    
    # Step 2: Extract contents into a temporary directory
    temp_dir = "temp_extract"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)  # Remove if exists
    os.makedirs(temp_dir)
    
    with zipfile.ZipFile(zip_file, 'r') as zf:
        zf.extractall(temp_dir)
    
    # Step 3: Modify workbook.xml to remove workbookProtection
    workbook_xml_path = os.path.join(temp_dir, "xl", "workbook.xml")
    if os.path.exists(workbook_xml_path):
        with open(workbook_xml_path, "r", encoding="utf-8") as file:
            content = file.read()
        
        # Remove the workbookProtection tag
        content = re.sub(r'<workbookProtection[^>]*>', '', content)
        
        with open(workbook_xml_path, "w", encoding="utf-8") as file:
            file.write(content)
    
    # Step 4: Recompress files back into the original .zip
    with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                zf.write(file_path, os.path.relpath(file_path, temp_dir))
    
    # Step 5: Rename .zip back to .xlsx (original file)
    os.rename(zip_file, excel_file)
    
    # Cleanup: Remove temporary extraction folder
    shutil.rmtree(temp_dir)
    
    print(f"Workbook protection removed in the original file: {excel_file}")

def unprotect_sheets(excel_file):
    # Load the workbook
    wb = openpyxl.load_workbook(excel_file)
    
    # Loop through all sheets and remove protection
    for sheet in wb.worksheets:
        try:
            sheet.protection.sheet = False  # Disable sheet protection
            sheet.protection.password = None  # Remove password if present
            print(f"Unprotected: {sheet.title}")
        except Exception as e:
            print(f"Failed to unprotect {sheet.title}: {e}")
    
    # Save changes directly to the same file
    wb.save(excel_file)
    print(f"All sheets unprotected in the original file: {excel_file}")

# Example Usage
protected_excel = r"C:\Users\dboliveira\Planilha Estimativas BdB216020 - Filtragem 3.xlsx"  # Change to your encrypted file

if __name__ == '__main__':
    try:
        unprotect_workbook(protected_excel)
        unprotect_sheets(protected_excel)
    except BaseException:
        import sys
        print(sys.exc_info()[0])
        import traceback
        print(traceback.format_exc())
    finally:
        print("Press Enter to continue ...")
        input()
