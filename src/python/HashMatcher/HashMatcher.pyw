import hashlib
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

# pip install pyinstaller ttkbootstrap

# Execute no terminal
# cd "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\0-GERENCIAMENTO\Script\Python\HashMatcher\"
# python -m PyInstaller --noconfirm --onefile --icon="HashMatcher.ico" --windowed --distpath "." "HashMatcher.pyw"

def show_help():
    help_text = (
        "🔐 HashMatcher – Comparador SHA-256\n\n"
        "Este programa permite comparar dois arquivos usando o algoritmo de hash SHA-256.\n\n"
        "✔ Use o botão 'Procurar' para selecionar dois arquivos.\n"
        "✔ Clique em 'Comparar Arquivos' para verificar se são idênticos em conteúdo.\n\n"
        "Por que usar?\n"
        "- Para validar integridade de arquivos.\n"
        "- Para saber se dois arquivos são exatamente iguais, byte a byte.\n"
        "- Para detectar alterações, duplicações ou corrupção de arquivos.\n\n"
        "SHA-256 é um algoritmo de hash criptográfico amplamente usado e seguro.\n\n"
        "Desenvolvido por Daniel Oliveira (danielbo17@hotmail.com)."
    )
    messagebox.showinfo("Ajuda - Sobre o Programa", help_text)

def calculate_sha256(file_path):
    sha256_hash = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        return f"Error: {str(e)}"

def browse_file(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, "end")
    entry.insert(0, file_path)

def compare_files():
    file1 = entry1.get()
    file2 = entry2.get()

    if not os.path.isfile(file1) or not os.path.isfile(file2):
        messagebox.showerror("Erro", "Selecione dois arquivos válidos.")
        return

    hash1 = calculate_sha256(file1)
    hash2 = calculate_sha256(file2)

    hash1_label.config(text=f"Arquivo 1 SHA-256:\n{hash1}")
    hash2_label.config(text=f"Arquivo 2 SHA-256:\n{hash2}")

    if hash1 == hash2:
        result_label.config(text="Resultado: Arquivos são IDÊNTICOS", foreground="green")
    else:
        result_label.config(text="Resultado: Arquivos são DIFERENTES", foreground="red")

app = ttk.Window(themename="superhero")  # Outros temas: "flatly", "darkly", "cosmo", "journal"
app.title("HashMatcher – Comparador SHA-256")
app.geometry("700x500")

ttk.Label(app, text="Arquivo 1:", font=("Segoe UI", 10)).pack(pady=(10, 0))
entry1 = ttk.Entry(app, width=90)
entry1.pack(pady=5)
ttk.Button(app, text="Procurar", command=lambda: browse_file(entry1)).pack(pady=5)

ttk.Label(app, text="Arquivo 2:", font=("Segoe UI", 10)).pack(pady=(10, 0))
entry2 = ttk.Entry(app, width=90)
entry2.pack(pady=5)
ttk.Button(app, text="Procurar", command=lambda: browse_file(entry2)).pack(pady=5)

ttk.Button(app, text="Comparar Arquivos", command=compare_files, bootstyle=PRIMARY).pack(pady=15)

# Botão de ajuda no topo
ttk.Button(app, text="Ajuda", command=show_help, bootstyle=INFO).place(x=10, y=10)

hash1_label = ttk.Label(app, text="", font=("Courier New", 9), wraplength=680, justify="left")
hash1_label.pack(pady=5)

hash2_label = ttk.Label(app, text="", font=("Courier New", 9), wraplength=680, justify="left")
hash2_label.pack(pady=5)

result_label = ttk.Label(app, text="", font=("Segoe UI", 12, "bold"))
result_label.pack(pady=15)

app.mainloop()
