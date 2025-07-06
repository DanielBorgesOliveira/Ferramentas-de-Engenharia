import os
import sys
import pypdf
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# pip install pyinstaller pypdf tkinter

# Execute no terminal
# cd "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\0-GERENCIAMENTO\Script\Python\AttachPDF"
# python -m PyInstaller --noconfirm --onefile --icon="AttachPDF.ico" --windowed --distpath "." "AttachPDF.pyw"

def _(key, lang, texts):
    return texts.get(key, {}).get(lang, key)


class PDFAttachmentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # Supported languages and mappings
        self.languages = ['English', 'Português', 'Español']
        self.lang_map = {'English': 'en', 'Português': 'pt', 'Español': 'es'}
        # Detect system locale
        try:
            sys_loc = locale.getdefaultlocale()[0] or ''
            lang_code = sys_loc.split('_')[0]
        except Exception:
            lang_code = 'en'
        if lang_code not in ('en', 'pt', 'es'):
            lang_code = 'en'
        self.current_lang = lang_code

        # Text translations
        self.texts = {
            'title': {'en': 'PDF Attachment Tool', 'pt': 'Ferramenta de Anexar PDF', 'es': 'Herramienta de Adjuntar PDF'},
            'input_pdf': {'en': 'Input PDF:', 'pt': 'PDF de entrada:', 'es': 'PDF de entrada:'},
            'browse': {'en': 'Browse', 'pt': 'Procurar', 'es': 'Buscar'},
            'attachments': {'en': 'Attachments:', 'pt': 'Arquivos anexos:', 'es': 'Archivos adjuntos:'},
            'add_files': {'en': 'Add Files', 'pt': 'Adicionar arquivos', 'es': 'Agregar archivos'},
            'remove_selected': {'en': 'Remove Selected', 'pt': 'Remover selecionados', 'es': 'Quitar seleccionados'},
            'clear_all': {'en': 'Clear All', 'pt': 'Limpar tudo', 'es': 'Limpiar todo'},
            'overwrite': {'en': 'Overwrite original', 'pt': 'Sobrescrever original', 'es': 'Sobrescribir original'},
            'attach_save': {'en': 'Attach and Save', 'pt': 'Anexar e Salvar', 'es': 'Adjuntar y Guardar'},
            'err_no_input': {'en': 'Please select an input PDF file.', 'pt': 'Por favor, selecione um arquivo PDF de entrada.', 'es': 'Por favor, seleccione un PDF de entrada.'},
            'err_no_attach': {'en': 'Please add at least one attachment.', 'pt': 'Por favor, adicione pelo menos um arquivo.', 'es': 'Por favor, agregue al menos un archivo.'},
            'success_saved': {'en': 'PDF saved:\n{path}', 'pt': 'PDF salvo:\n{path}', 'es': 'PDF guardado:\n{path}'},
        }

        self.geometry("600x450")
        self.resizable(False, False)
        self.input_file = None
        self.attached_files = []
        self.create_widgets()
        # Set combobox to system language
        lang_name = next((name for name, code in self.lang_map.items() if code == self.current_lang), 'English')
        self.combo_lang.current(self.languages.index(lang_name))
        self.update_language()

    def create_widgets(self):
        style = ttk.Style(self)
        style.theme_use('clam')

        # Language selection
        lang_frame = ttk.Frame(self, padding=10)
        lang_frame.pack(fill='x')
        ttk.Label(lang_frame, text='Language:').pack(side='left')
        self.combo_lang = ttk.Combobox(lang_frame, values=self.languages, state='readonly', width=12)
        self.combo_lang.pack(side='left', padx=(5, 0))
        self.combo_lang.bind('<<ComboboxSelected>>', lambda e: self.on_language_change())

        # Input PDF selection
        input_frame = ttk.Frame(self, padding=10)
        input_frame.pack(fill='x')
        self.label_input = ttk.Label(input_frame)
        self.label_input.pack(side='left', padx=(0, 5))
        self.input_entry = ttk.Entry(input_frame, width=50)
        self.input_entry.pack(side='left', fill='x', expand=True)
        self.btn_browse = ttk.Button(input_frame, command=self.select_input)
        self.btn_browse.pack(side='left', padx=5)

        # Attachments list
        attach_frame = ttk.Frame(self, padding=10)
        attach_frame.pack(fill='both', expand=True)
        self.label_attach = ttk.Label(attach_frame)
        self.label_attach.pack(anchor='nw')
        self.attach_listbox = tk.Listbox(attach_frame, height=10)
        self.attach_listbox.pack(fill='both', expand=True, pady=5)

        btn_frame = ttk.Frame(attach_frame)
        btn_frame.pack(fill='x')
        self.btn_add = ttk.Button(btn_frame, command=self.add_attachments)
        self.btn_add.pack(side='left')
        self.btn_remove = ttk.Button(btn_frame, command=self.remove_selected)
        self.btn_remove.pack(side='left', padx=5)
        self.btn_clear = ttk.Button(btn_frame, command=self.clear_all)
        self.btn_clear.pack(side='left')

        # Overwrite option and process button
        process_frame = ttk.Frame(self, padding=10)
        process_frame.pack(fill='x')
        self.overwrite_var = tk.BooleanVar(value=False)
        self.chk_overwrite = ttk.Checkbutton(process_frame, variable=self.overwrite_var)
        self.chk_overwrite.pack(side='left', padx=(0, 10))
        self.btn_process = ttk.Button(process_frame, command=self.process)
        self.btn_process.pack(side='left')

    def on_language_change(self):
        lang_name = self.combo_lang.get()
        self.current_lang = self.lang_map.get(lang_name, 'en')
        self.update_language()

    def update_language(self):
        lang = self.current_lang
        self.title(_('title', lang, self.texts))
        self.label_input.config(text=_('input_pdf', lang, self.texts))
        self.btn_browse.config(text=_('browse', lang, self.texts))
        self.label_attach.config(text=_('attachments', lang, self.texts))
        self.btn_add.config(text=_('add_files', lang, self.texts))
        self.btn_remove.config(text=_('remove_selected', lang, self.texts))
        self.btn_clear.config(text=_('clear_all', lang, self.texts))
        self.chk_overwrite.config(text=_('overwrite', lang, self.texts))
        self.btn_process.config(text=_('attach_save', lang, self.texts))

    def select_input(self):
        filetypes = [(_('input_pdf', self.current_lang, self.texts), "*.pdf")]
        filepath = filedialog.askopenfilename(title=_('browse', self.current_lang, self.texts), filetypes=filetypes)
        if filepath:
            self.input_file = filepath
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filepath)

    def add_attachments(self):
        files = filedialog.askopenfilenames(title=_('add_files', self.current_lang, self.texts))
        for f in files:
            if f not in self.attached_files:
                self.attached_files.append(f)
                self.attach_listbox.insert(tk.END, f)

    def remove_selected(self):
        for index in reversed(list(self.attach_listbox.curselection())):
            self.attached_files.pop(index)
            self.attach_listbox.delete(index)

    def clear_all(self):
        self.attached_files.clear()
        self.attach_listbox.delete(0, tk.END)

    def process(self):
        if not self.input_file:
            messagebox.showerror(_('title', self.current_lang, self.texts), _('err_no_input', self.current_lang, self.texts))
            return
        if not self.attached_files:
            messagebox.showerror(_('title', self.current_lang, self.texts), _('err_no_attach', self.current_lang, self.texts))
            return
        try:
            reader = pypdf.PdfReader(self.input_file)
            writer = pypdf.PdfWriter()
            writer.append_pages_from_reader(reader)
            for filepath in self.attached_files:
                name = os.path.basename(filepath)
                with open(filepath, "rb") as f:
                    writer.add_attachment(name, f.read())
            if self.overwrite_var.get():
                save_path = self.input_file
            else:
                save_path = filedialog.asksaveasfilename(title=_('attach_save', self.current_lang, self.texts), defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if not save_path:
                return
            with open(save_path, "wb") as out_f:
                writer.write(out_f)
            messagebox.showinfo(_('title', self.current_lang, self.texts), _('success_saved', self.current_lang, self.texts).format(path=save_path))
        except Exception as e:
            messagebox.showerror(_('title', self.current_lang, self.texts), str(e))


if __name__ == "__main__":
    app = PDFAttachmentApp()
    app.mainloop()