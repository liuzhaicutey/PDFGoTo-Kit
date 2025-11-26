import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz
import os
import sys
import platform
from PIL import Image, ImageTk
import math
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import Color
import io
import tempfile
import glob
from tkinter import font as tkFont
import time
from docx2pdf import convert

angle = 45
m = fitz.Matrix(angle)

m2 = fitz.Matrix(1, 1)
m2.prerotate(angle)

AUTHOR_NAME = "Liu Zhai"
SOFTWARE_VERSION = "1.2.0"
BUILD_ID = "PDFGoToKit-20251125-XYZ789"

WINDOWS_ICON_FILENAME = 'app_icon.ico'
NON_WINDOWS_ICON_FILENAME = 'app_icon.png'
THEME_FILE_NAME = "pdfgotokit_theme.txt" 
THEME_FILE_PATH = os.path.join(tempfile.gettempdir(), THEME_FILE_NAME)

THEMES = {
    "Blue": {
        'PRIMARY': '#1976D2',
        'ACCENT': '#1565C0',
        'LIGHT_BG': '#FFFFFF',
        'MID_GRAY': '#EEEEEE',
        'DARK_TEXT': '#212121',
        'ERROR': '#A00000'
    },
    "Charcoal": {
        'PRIMARY': '#00ADB5',
        'ACCENT': '#008388',
        'LIGHT_BG': '#393E46',
        'MID_GRAY': '#505862',
        'DARK_TEXT': '#EEEEEE',
        'ERROR': '#FF6B6B'
    },
    "Forest": {
        'PRIMARY': '#388E3C',
        'ACCENT': '#2E7D32',
        'LIGHT_BG': '#F0F4C3',
        'MID_GRAY': '#E6EE9C',
        'DARK_TEXT': '#1B5E20',
        'ERROR': '#D32F2F'
    },
    "Sunset": {
        'PRIMARY': '#FF9800',
        'ACCENT': '#F57C00',
        'LIGHT_BG': '#FFF3E0',
        'MID_GRAY': '#FFCC80',
        'DARK_TEXT': '#E65100',
        'ERROR': '#C62828'
    },
    "Violet": {
        'PRIMARY': '#673AB7',
        'ACCENT': '#5E35B1',
        'LIGHT_BG': '#EDE7F6',
        'MID_GRAY': '#D1C4E9',
        'DARK_TEXT': '#311B92',
        'ERROR': '#B71C1C'
    }
}

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class PDFToolkitApp:
    
    BASE14 = [
        "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
        "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
        "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique",
        "Symbol", "ZapfDingbats"
    ]

    def __init__(self, master):
        self.master = master
        master.title(f"PDFGoTo Kit by {AUTHOR_NAME}")
        master.geometry("680x600")
        master.resizable(False, False)
        initial_theme = self._load_saved_theme()
        self.current_theme = tk.StringVar(value=initial_theme)
        self.colors = THEMES[self.current_theme.get()]
        self.merge_tab = ttk.Frame(master)

        self.merge_file_paths = []
        self._setup_merge_tab()

        self.apply_theme()
        

        self.current_preview_font = None


        try:
            if platform.system() == "Windows":
                master.iconbitmap(resource_path(WINDOWS_ICON_FILENAME))
            else:
                img = Image.open(resource_path(NON_WINDOWS_ICON_FILENAME))
                self.icon_img = ImageTk.PhotoImage(img.resize((32, 32)))
                master.iconphoto(True, self.icon_img)
        except Exception:
            pass

        self.main_frame = ttk.Frame(master, padding="20")
        self.main_frame.pack(fill="both", expand=True)

        ttk.Label(self.main_frame, text="📄 PDFGoTo Kit", style='Title.TLabel').grid(row=0, column=0, sticky='w', pady=(0, 10))
        ttk.Label(self.main_frame, text=f"© 2025 {AUTHOR_NAME}. Version {SOFTWARE_VERSION}", font=('Arial', 8, 'italic'), foreground='#555555').grid(row=0, column=1, sticky='e', pady=(0,10))

        self.theme_selector_frame = ttk.Frame(self.main_frame, padding="10 5")
        self.theme_selector_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0,5))
        self._setup_theme_selector()

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(5,5), padx=5)

        self.merge_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.merge_tab, text='Merge')
        self._setup_merge_tab()

        self.split_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.split_tab, text='Split')
        self._setup_split_tab()

        self.convert_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.convert_tab, text='Convert')
        self._setup_convert_tab()


        self.rotate_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.rotate_tab, text='Rotate')
        self._setup_rotate_tab()

        self.watermark_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.watermark_tab, text='Watermark')
        self._setup_watermark_tab()

        self.stamp_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.stamp_tab, text='Stamp')
        self._setup_stamp_tab()
        
        self.encrypt_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.encrypt_tab, text='Encrypt/Decrypt')
        self._setup_encrypt_tab()
        
        self.metadata_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.metadata_tab, text='Metadata')
        self._setup_metadata_tab()

        self.footer_label = ttk.Label(
            master, 
            text="This software is provided free for distribution. Resale for profit is strictly prohibited.",
            font=('Arial', 8, 'italic'), 
            foreground=self.colors['ERROR'],
            background=self.colors['LIGHT_BG']
        )
        self.footer_label.pack(pady=(5,10))

    def change_theme(self, theme_name):
        self.colors = THEMES[theme_name]
        self.apply_theme()
        self._save_current_theme(theme_name)

    def apply_theme(self):
        self.master.configure(bg=self.colors['LIGHT_BG'])
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass

        style.configure('TFrame', background=self.colors['LIGHT_BG'])
        style.configure('TLabel', background=self.colors['LIGHT_BG'], foreground=self.colors['DARK_TEXT'], font=('Arial', 10))
        style.configure('TNotebook', background=self.colors['LIGHT_BG'], borderwidth=0)
        style.map('TNotebook.Tab', background=[('selected', self.colors['MID_GRAY'])])
        style.configure('TNotebook.Tab', padding=[10, 5], background=self.colors['MID_GRAY'], foreground=self.colors['DARK_TEXT'], font=('Arial', 10, 'bold'))
        style.configure('TEntry', fieldbackground=self.colors['MID_GRAY'], foreground=self.colors['DARK_TEXT'], borderwidth=1)
        style.map('TEntry', fieldbackground=[('readonly', self.colors['MID_GRAY'])])
        style.configure('TButton', font=('Arial', 10, 'bold'), padding=8, foreground='white', background=self.colors['PRIMARY'], borderwidth=0, relief='flat')
        style.map('TButton', background=[('active', self.colors['ACCENT'])], foreground=[('active', 'white')])
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground=self.colors['PRIMARY'], background=self.colors['LIGHT_BG'])
        style.configure('Accent.TButton', background=self.colors['PRIMARY'], foreground='white', font=('Arial', 11, 'bold'), padding=10)
        style.map('Accent.TButton', background=[('active', self.colors['ACCENT'])])

        if hasattr(self, 'footer_label'):
            self.footer_label.configure(background=self.colors['LIGHT_BG'], foreground=self.colors['ERROR'])

        if hasattr(self, 'merge_listbox'):
            self.merge_listbox.configure(bg=self.colors['MID_GRAY'], fg=self.colors['DARK_TEXT'])

        if hasattr(self, 'theme_selector_frame'):
             self.theme_selector_frame.configure(style='TFrame')

    def _setup_theme_selector(self):
        selector_frame = ttk.Frame(self.theme_selector_frame, style='TFrame')
        selector_frame.pack(fill='x')

        ttk.Label(selector_frame, text="Current Theme:", font=('Arial', 10, 'bold')).pack(side='left', padx=(0, 10))

        theme_names = list(THEMES.keys())

        style = ttk.Style()
        style.configure('ThemeLabel.TLabel', background=self.colors['LIGHT_BG'], foreground=self.colors['DARK_TEXT'], font=('Arial', 10, 'bold'))

        theme_combo = ttk.Combobox(selector_frame,
                                       textvariable=self.current_theme,
                                       values=theme_names,
                                       state='readonly',
                                       width=20)

        theme_combo.bind("<<ComboboxSelected>>", lambda event: self.change_theme(self.current_theme.get()))
        theme_combo.pack(side='left', fill='x', expand=False)
        theme_combo.set(self.current_theme.get())

    def _load_saved_theme(self):
        try:
            with open(THEME_FILE_PATH, 'r') as f:
                saved_theme = f.read().strip()
                if saved_theme in THEMES:
                    return saved_theme
        except FileNotFoundError:

            pass 
        except Exception:

            pass 
        return "Blue"
    
    def _save_current_theme(self, theme_name):

        try:
            with open(THEME_FILE_PATH, 'w') as f:
                f.write(theme_name)
        except Exception as e:

            print(f"Error saving theme: {e}")

    def _setup_merge_tab(self):
        ttk.Label(self.merge_tab, text="Files to Merge (Order Matters):", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor='w')

        list_frame = ttk.Frame(self.merge_tab)
        list_frame.pack(fill='x', expand=True, padx=0, pady=5)

        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)

        self.merge_listbox = tk.Listbox(list_frame, height=15, width=50, selectmode=tk.SINGLE,
                                        yscrollcommand=list_scrollbar.set, bg=self.colors['MID_GRAY'], fg=self.colors['DARK_TEXT'], bd=0, highlightthickness=0)
        list_scrollbar.config(command=self.merge_listbox.yview)

        list_scrollbar.pack(side='right', fill='y')
        self.merge_listbox.pack(side='left', fill='both', expand=True)


        self.merge_listbox.bind('<Button-1>', self._start_drag)
        self.merge_listbox.bind('<B1-Motion>', self._drag_motion)
        self.merge_listbox.bind('<ButtonRelease-1>', self._drop_file)


        btn_frame = ttk.Frame(self.merge_tab, style='TFrame')
        btn_frame.pack(fill='x', pady=10)

        ttk.Button(btn_frame, text="➕ Add PDF", command=self.add_merge_file).pack(side='left', padx=(0, 10), fill='x', expand=True)
        ttk.Button(btn_frame, text="❌ Remove Selected", command=self.remove_merge_file).pack(side='left', fill='x', expand=True)

        ttk.Button(self.merge_tab, text="🚀 Merge PDFs", style='Accent.TButton', command=self.merge_pdfs).pack(fill='x', pady=(10, 0))

        self.dragged_item_index = None
        self.dragged_item_path = None 
        self.merge_listbox.bind('<Up>', lambda e: self._move_selected_file(-1))
        self.merge_listbox.bind('<Down>', lambda e: self._move_selected_file(1))

    def _update_listbox_display(self):

        self.merge_listbox.delete(0, tk.END)
        for i, file_path in enumerate(self.merge_file_paths):

            file_name = os.path.basename(file_path)
            

            numbered_display = f" {i + 1: >2}.  {file_name}"
            self.merge_listbox.insert(tk.END, numbered_display)

    def _start_drag(self, event):
        try:
            index = self.merge_listbox.nearest(event.y)
            if index in self.merge_listbox.curselection() or self.merge_listbox.size() > 0:
                self.dragged_item_index = index

                self.dragged_item_path = self.merge_file_paths[index]
                self.merge_listbox.config(cursor='hand2')
        except IndexError:
            self.dragged_item_index = None

    def _drag_motion(self, event):
        if self.dragged_item_index is not None:
            target_index = self.merge_listbox.nearest(event.y)

            if target_index < 0 or target_index >= len(self.merge_file_paths):
                return

            if target_index != self.dragged_item_index:

                self.merge_file_paths.pop(self.dragged_item_index)
                self.merge_file_paths.insert(target_index, self.dragged_item_path)

                self._update_listbox_display()

                self.dragged_item_index = target_index
                self.merge_listbox.selection_set(target_index)

    def _move_selected_file(self, direction):

        try:

            selected_indices = self.merge_listbox.curselection()
            if not selected_indices:
                return 

            current_index = selected_indices[0]
            new_index = current_index + direction 
            list_size = len(self.merge_file_paths)

            if 0 <= new_index < list_size:

                

                file_path_to_move = self.merge_file_paths.pop(current_index)
                

                self.merge_file_paths.insert(new_index, file_path_to_move)

                self._update_listbox_display()


                self.merge_listbox.selection_clear(0, tk.END)
                self.merge_listbox.selection_set(new_index)
                self.merge_listbox.activate(new_index)
                self.merge_listbox.see(new_index)

        except Exception as e:

            print(f"Error moving file: {e}")

    def _drop_file(self, event):

        self.dragged_item_index = None
        self.dragged_item_path = None
        self.merge_listbox.config(cursor='')

    def _setup_split_tab(self):
        self.split_filepath = tk.StringVar()

        ttk.Label(self.split_tab, text="Selected PDF:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor='w')
        filepath_entry = ttk.Entry(self.split_tab, textvariable=self.split_filepath, state='readonly', width=50)
        filepath_entry.pack(fill='x', padx=0, pady=(0, 10))

        ttk.Button(self.split_tab, text="📂 Select PDF to Split", command=self.select_split_file).pack(fill='x', pady=5)

        ttk.Label(self.split_tab, text="How to Split:", font=('Arial', 10, 'bold')).pack(pady=(15, 5), anchor='w')
        ttk.Button(self.split_tab, text="✂️ Split into Single Pages (1.pdf, 2.pdf, etc.)", command=self.split_pdf_all).pack(fill='x', pady=5)

        ttk.Label(self.split_tab, text="Or, Split a Range (e.g., 5-10):", font=('Arial', 10, 'bold')).pack(pady=(15, 5), anchor='w')
        range_frame = ttk.Frame(self.split_tab)
        range_frame.pack(fill='x')
        self.split_range = tk.StringVar()
        ttk.Entry(range_frame, textvariable=self.split_range, width=15).pack(side='left', fill='x', padx=(0, 10), expand=True)
        ttk.Button(range_frame, text="✂️ Split Custom Range", style='Accent.TButton', command=self.split_pdf_range).pack(side='left', fill='x', expand=True)

    def _setup_convert_tab(self):
        self.convert_file_list = []

        ttk.Label(
            self.convert_tab,
            text="Conversion Tool",
            font=('Arial', 15, 'bold')
        ).pack(pady=(0, 2), anchor='w')


        ttk.Label(
            self.convert_tab,
            text="Note: Some conversions require Microsoft Office or LibreOffice installed on this system.",
            foreground="#777777",
            font=('Arial', 9, 'italic')
        ).pack(pady=(0, 10), anchor='w')


        list_frame = ttk.Frame(self.convert_tab)
        list_frame.pack(fill='x', expand=True, pady=5)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.convert_listbox = tk.Listbox(
            list_frame,
            height=10,
            width=50,
            selectmode=tk.SINGLE,
            yscrollcommand=scrollbar.set,
            bg=self.colors['MID_GRAY'],
            fg=self.colors['DARK_TEXT'],
            bd=0,
            highlightthickness=0
        )
        scrollbar.config(command=self.convert_listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.convert_listbox.pack(side='left', fill='both', expand=True)

        btn_frame = ttk.Frame(self.convert_tab)
        btn_frame.pack(fill='x', pady=10)

        ttk.Button(btn_frame, text="➕ Add Files", command=self.add_convert_files).pack(
            side='left', padx=(0, 10), expand=True, fill='x'
        )
        ttk.Button(btn_frame, text="❌ Remove Selected", command=self.remove_convert_file).pack(
            side='left', expand=True, fill='x'
        )

        ttk.Label(
            self.convert_tab,
            text="Select Conversion Type:",
            font=('Arial', 10, 'bold')
        ).pack(pady=(5, 2), anchor='w')

        self.conversion_type = tk.StringVar()
        self.convert_combo = ttk.Combobox(
            self.convert_tab,
            textvariable=self.conversion_type,
            state="readonly",
            width=40,
            values=[
                "PDF → TXT",
                "PDF → DOCX",
                "PDF → XLSX",
                "DOCX → PDF",
                "XLSX → PDF",
                "Images → PDF"
            ]
        )
        self.convert_combo.pack(anchor='w', pady=(0, 5))

        self.extension_error_msg = tk.StringVar(value="")
        self.error_label = ttk.Label(
            self.convert_tab,
            textvariable=self.extension_error_msg,
            foreground="red", 
            font=('Arial', 9, 'bold')
        )
        self.error_label.pack(anchor='w', pady=(0, 5))

        ttk.Button(
            self.convert_tab,
            text="🚀 Convert",
            style='Accent.TButton',
            command=self.perform_conversion
        ).pack(fill='x', pady=(10, 0))

    def _get_extension(self, filepath):

        return os.path.splitext(filepath)[1].lower()

    def _check_and_display_extension_error(self):

        if not self.convert_file_list:

            self.extension_error_msg.set("")
            return True


        extensions = {self._get_extension(f) for f in self.convert_file_list}
        

        image_extensions = {".jpg", ".jpeg", ".png"}
        

        if extensions.issubset(image_extensions) and extensions:
            self.extension_error_msg.set("")
            return True


        if len(extensions) > 1:

            if not extensions.issubset(image_extensions):
                self.extension_error_msg.set("⚠️ All files must share the same primary extension (e.g., all PDF, all DOCX).")

                return False
                
        self.extension_error_msg.set("")
        return True


    def update_convert_dropdown(self):

        if not self.convert_file_list:
            self.convert_combo.config(values=[], state="disabled")
            self.conversion_type.set("")
            return


        exts = {os.path.splitext(f)[1].lower() for f in self.convert_file_list}


        allowed_modes = []


        if exts == {".pdf"}:
            allowed_modes = [
                "PDF → TXT",
                "PDF → DOCX",
                "PDF → XLSX"
            ]

        elif exts == {".docx"}:
            allowed_modes = [
                "DOCX → PDF"
            ]


        elif exts == {".xlsx"}:
            allowed_modes = [
                "XLSX → PDF"
            ]


        elif exts.issubset({".jpg", ".jpeg", ".png"}):
            allowed_modes = [
                "Images → PDF"
            ]


        else:
            allowed_modes = []
            self.convert_combo.config(state="disabled")
            self.conversion_type.set("")
            return


        self.convert_combo.config(state="readonly", values=allowed_modes)


        if allowed_modes:
            self.conversion_type.set(allowed_modes[0])

    def add_convert_files(self):

        if not self.convert_file_list:
            filetypes = [
                ("All Supported", "*.pdf *.docx *.xlsx *.jpg *.jpeg *.png"),
                ("PDF files", "*.pdf"),
                ("Word documents", "*.docx"),
                ("Excel files", "*.xlsx"),
                ("Images", "*.jpg *.jpeg *.png")
            ]
        else:
            exts = {os.path.splitext(f)[1].lower() for f in self.convert_file_list}

            if exts == {".pdf"}:
                filetypes = [("PDF files", "*.pdf")]
            elif exts == {".docx"}:
                filetypes = [("Word documents", "*.docx")]
            elif exts == {".xlsx"}:
                filetypes = [("Excel files", "*.xlsx")]
            elif exts.issubset({".jpg", ".jpeg", ".png"}):
                filetypes = [("Images", "*.jpg *.jpeg *.png")]
            else:
                messagebox.showerror("Error", "Invalid mixed file types. Clear the list first.")
                return

        new_files = filedialog.askopenfilenames(title="Select Files", filetypes=filetypes)
        if not new_files:
            return

        current_list_size = len(self.convert_file_list)

        for f in new_files:
            self.convert_file_list.append(f)
            file_name = os.path.basename(f)
            self.convert_listbox.insert(tk.END, file_name)

            # 🔥 FIX for DOCX → PDF converter
            ext = os.path.splitext(f)[1].lower()
            if ext == ".docx":
                self._docx_file_full_path = f  # THIS LINE FIXES THE ERROR

        self.update_convert_dropdown()

        if new_files:
            first_new_index = current_list_size
            self.convert_listbox.selection_clear(0, tk.END)
            self.convert_listbox.selection_set(first_new_index)
            self.convert_listbox.see(first_new_index)


    def remove_convert_file(self):
        sel = self.convert_listbox.curselection()
        if not sel:
            return
        index = sel[0]

        self.convert_listbox.delete(index)
        del self.convert_file_list[index]

        self.update_convert_dropdown()

    def repair_pdf(self, filepath):
        try:
            doc = fitz.open(filepath)

            pdf_bytes = doc.tobytes(clean=True, garbage=4, deflate=True)
            doc.close()

            return pdf_bytes
        except Exception:

            with open(filepath, "rb") as f:
                return f.read()

    def perform_conversion(self):
        selection = self.conversion_type.get().strip()

        if not selection:
            messagebox.showwarning("Warning", "Please choose a conversion type.")
            return

        if not self.convert_file_list:
            messagebox.showwarning("Warning", "Please add files first.")
            return

        if selection == "PDF → TXT":
            self.convert_pdf_to_txt()

        elif selection == "PDF → DOCX":
            self.convert_pdf_to_docx()

        elif selection == "PDF → XLSX":
            self.convert_pdf_to_xlsx()

        elif selection == "DOCX → PDF":
            self.convert_docx_to_pdf()

        elif selection == "XLSX → PDF":
            self.convert_xlsx_to_pdf()

        elif selection == "Images → PDF":
            self.convert_images_to_pdf()

        else:
            messagebox.showerror("Error", "Unknown conversion type.")

    def convert_docx_to_pdf(self):
        if not hasattr(self, "_docx_file_full_path") or not self._docx_file_full_path:
            messagebox.showwarning("Warning", "Please select a DOCX file first.")
            return

        input_path = self._docx_file_full_path

        output_pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Converted PDF As"
        )

        if not output_pdf_path:
            return 

        try:
            convert(input_path, output_pdf_path)

            if os.path.exists(output_pdf_path):
                messagebox.showinfo("Success", f"Successfully converted to PDF:\n{output_pdf_path}")
            else:
                messagebox.showerror("Error", "Conversion failed: output file was not created.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert DOCX to PDF:\n{e}")

    def convert_pdf_to_docx(self):
        try:
            from pdf2docx import Converter
        except:
            messagebox.showerror("Error", "Missing module: pdf2docx")
            return

        output_dir = filedialog.askdirectory(title="Select Folder to Save DOCX Files")
        if not output_dir:
            return

        for f in self.convert_file_list:

            repaired_bytes = self.repair_pdf(f)

            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            temp_pdf.write(repaired_bytes)
            temp_pdf.close()

            out = os.path.join(
                output_dir,
                os.path.splitext(os.path.basename(f))[0] + ".docx"
            )
            name = os.path.splitext(os.path.basename(f))[0] + ".docx"
            out = os.path.join(output_dir, name)

            try:
                cv = Converter(f)
                cv.convert(out)
                cv.close()
            except Exception as e:
                messagebox.showerror("Error", f"Failed converting {f}\n{e}")
                return

        messagebox.showinfo("Success", "All PDFs converted to DOCX!")

    def convert_pdf_to_xlsx(self):
        try:
            import camelot
            import pandas as pd
            import os 
            import tempfile
            from tkinter import filedialog, messagebox 
        except ImportError:
            messagebox.showerror("Error", "Missing modules", "The 'camelot' or 'pandas' modules are not installed. Please install them to use this feature.")
            return

        output_dir = filedialog.askdirectory(title="Save XLSX Files To")
        if not output_dir:
            return

        total_files = len(self.convert_file_list)
        successful_conversions = 0

        for f in self.convert_file_list:
            temp_pdf_path = None
            file_base_name = os.path.basename(f)

            try:

                repaired_bytes = self.repair_pdf(f)

                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                temp_pdf.write(repaired_bytes)
                temp_pdf.close()
                temp_pdf_path = temp_pdf.name 


                tables_lattice = camelot.read_pdf(
                    temp_pdf_path, 
                    pages="all", 
                    flavor='lattice'
                )
                tables_stream = camelot.read_pdf(
                    temp_pdf_path, 
                    pages="all", 
                    flavor='stream'
                )
                

                tables = list(tables_lattice) + list(tables_stream)
                

                if len(tables) == 0:
                    messagebox.showwarning("No Tables", f"No extractable tables found in {file_base_name} after trying both 'lattice' and 'stream' methods.")
                    continue


                out_xlsx_path = os.path.join(
                    output_dir,
                    os.path.splitext(file_base_name)[0] + ".xlsx"
                )


                writer = pd.ExcelWriter(out_xlsx_path, engine='xlsxwriter')
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
                writer.close()
                

                if os.path.exists(out_xlsx_path):
                    successful_conversions += 1
                else:
                    messagebox.showerror("Error", f"Failed to save XLSX for {file_base_name}: Output file was not created.")

            except Exception as e:

                messagebox.showerror("Error", f"Failed to process {file_base_name}:\n{e}")
                
            finally:

                if temp_pdf_path and os.path.exists(temp_pdf_path):
                    os.unlink(temp_pdf_path)


        if successful_conversions == total_files:
            messagebox.showinfo("Success", f"All {total_files} PDF files successfully converted to XLSX!")
        elif successful_conversions > 0:
            messagebox.showwarning("Partial Success", f"Successfully converted {successful_conversions} of {total_files} PDF files to XLSX. Check individual error messages for details on failed files.")
        else:
            messagebox.showerror("Failure", f"No tables were converted to XLSX. No tables were extracted from the selected files.")


    def convert_xlsx_to_pdf(self):
        import pandas as pd
        from reportlab.platypus import SimpleDocTemplate, Table
        from reportlab.lib.pagesizes import letter

        output_dir = filedialog.askdirectory(title="Save PDF Files To")
        if not output_dir:
            return

        for f in self.convert_file_list:
            df = pd.read_excel(f)
            data = [df.columns.tolist()] + df.values.tolist()

            out = os.path.join(
                output_dir,
                os.path.splitext(os.path.basename(f))[0] + ".pdf"
            )

            doc = SimpleDocTemplate(out, pagesize=letter)
            table = Table(data)
            doc.build([table])

        messagebox.showinfo("Success", "All XLSX files converted to PDF!")

    def convert_images_to_pdf(self):
        output_path = filedialog.asksaveasfilename(
            title="Save PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_path:
            return

        try:
            images = []
            for f in self.convert_file_list:
                img = Image.open(f).convert("RGB")
                images.append(img)

            images[0].save(output_path, save_all=True, append_images=images[1:])
            messagebox.showinfo("Success", "Images converted to a single PDF!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed converting images.\n{e}")


    def convert_pdf_to_txt(self):
        output_dir = filedialog.askdirectory(title="Select Folder to Save TXT Files")
        if not output_dir:
            return

        for f in self.convert_file_list:

            repaired_bytes = self.repair_pdf(f)

            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            temp_pdf.write(repaired_bytes)
            temp_pdf.close()

            out = os.path.join(
                output_dir,
                os.path.splitext(os.path.basename(f))[0] + ".docx"
            )

            try:
                doc = fitz.open(f)
                text = ""
                for p in doc:
                    text += p.get_text() + "\n\n"
                doc.close()

                out = os.path.join(output_dir, os.path.splitext(os.path.basename(f))[0] + ".txt")
                with open(out, "w", encoding="utf-8") as t:
                    t.write(text)
            except Exception as e:
                messagebox.showerror("Error", f"Failed extracting text from {f}\n{e}")
                return

        messagebox.showinfo("Success", "All PDFs converted to TXT!")


    def _setup_encrypt_tab(self):

        self.show_password_var = tk.BooleanVar()
        self.show_password_var.set(False) 
        self.show_password_var.trace_add("write", self.toggle_password_visibility)

        self.encrypt_filepath = tk.StringVar()
        self.encrypt_password = tk.StringVar()
        self.encrypt_confirm_password = tk.StringVar()
        self.decrypt_password = tk.StringVar()
        
        self.metadata_filepath = tk.StringVar()
        self.metadata_title = tk.StringVar()
        self.metadata_author = tk.StringVar()
        self.metadata_subject = tk.StringVar()
        self.metadata_creator = tk.StringVar()

        selector_frame = ttk.Frame(self.encrypt_tab)
        selector_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(selector_frame, text="🔒 Encrypt", command=lambda: self.show_encrypt_decrypt_frame('encrypt')).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(selector_frame, text="🔓 Decrypt", command=lambda: self.show_encrypt_decrypt_frame('decrypt')).pack(side='left', padx=5, fill='x', expand=True)
        
        ttk.Label(self.encrypt_tab, text="Selected PDF:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor='w')
        ttk.Entry(self.encrypt_tab, textvariable=self.encrypt_filepath, state='readonly', width=50).pack(fill='x', padx=0, pady=(0, 10))
        ttk.Button(self.encrypt_tab, text="📂 Select PDF", command=self.select_encrypt_file).pack(fill='x', pady=5)
        
        self.operation_frame = ttk.Frame(self.encrypt_tab)
        self.operation_frame.pack(fill='both', expand=True, pady=(15, 0))
        
        self.encrypt_frame = ttk.Frame(self.operation_frame)
        self.decrypt_frame = ttk.Frame(self.operation_frame)
        
        ttk.Label(self.encrypt_frame, text="Password for Encryption:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor='w')
        self.encrypt_password_entry = ttk.Entry(self.encrypt_frame, textvariable=self.encrypt_password, show='*', width=50)
        self.encrypt_password_entry.pack(fill='x', pady=5)

        ttk.Label(self.encrypt_frame, text="Confirm Password:", font=('Arial', 10, 'bold')).pack(pady=(10, 5), anchor='w')
        self.encrypt_confirm_password_entry = ttk.Entry(self.encrypt_frame, textvariable=self.encrypt_confirm_password, show='*', width=50)
        self.encrypt_confirm_password_entry.pack(fill='x', pady=5)
        
        ttk.Checkbutton(
            self.encrypt_frame, 
            text="Show Password", 
            variable=self.show_password_var
        ).pack(pady=(5, 5), anchor='w')

        ttk.Button(self.encrypt_frame, text="🔒 Encrypt PDF", style='Accent.TButton', command=self.encrypt_pdf).pack(fill='x', pady=(20, 0))

        ttk.Label(self.decrypt_frame, text="Password to Decrypt PDF:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor='w')
        self.decrypt_password_entry = ttk.Entry(self.decrypt_frame, textvariable=self.decrypt_password, show='*', width=50)
        self.decrypt_password_entry.pack(fill='x', pady=5)
        
        ttk.Label(self.decrypt_frame, text="Note: Decryption saves an unlocked copy.", foreground='#555555').pack(pady=(5, 10))

        ttk.Button(self.decrypt_frame, text="🔓 Decrypt PDF", style='Accent.TButton', command=self.decrypt_pdf).pack(fill='x', pady=(20, 0))
        

        self.show_encrypt_decrypt_frame('encrypt')

    def show_encrypt_decrypt_frame(self, mode):

        self.encrypt_frame.pack_forget()
        self.decrypt_frame.pack_forget()
        

        if mode == 'encrypt':
            self.encrypt_frame.pack(fill='both', expand=True)
        elif mode == 'decrypt':
            self.decrypt_frame.pack(fill='both', expand=True)

    def _setup_rotate_tab(self):
        self.rotate_filepath = tk.StringVar()
        ttk.Label(self.rotate_tab, text="Selected PDF:", font=('Arial',10,'bold')).pack(pady=(0,5), anchor='w')
        ttk.Entry(self.rotate_tab, textvariable=self.rotate_filepath, state='readonly', width=50).pack(fill='x', pady=(0,10))
        ttk.Button(self.rotate_tab, text="📂 Select PDF to Rotate", command=self.select_rotate_file).pack(fill='x', pady=5)
        ttk.Label(self.rotate_tab, text="Select Rotation Angle:", font=('Arial',10,'bold')).pack(pady=(15,5), anchor='w')
        angle_frame = ttk.Frame(self.rotate_tab)
        angle_frame.pack(fill='x', pady=5)
        ttk.Button(angle_frame, text="↻ 90°", command=lambda: self.rotate_pdf(90)).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(angle_frame, text="↻ 180°", command=lambda: self.rotate_pdf(180)).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(angle_frame, text="↻ 270°", command=lambda: self.rotate_pdf(270)).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Label(self.rotate_tab, text="Or specify custom angle:", font=('Arial',10,'bold')).pack(pady=(15,5), anchor='w')
        custom_frame = ttk.Frame(self.rotate_tab)
        custom_frame.pack(fill='x')
        self.rotate_angle = tk.StringVar(value="45")
        ttk.Entry(custom_frame, textvariable=self.rotate_angle, width=10).pack(side='left', padx=(0,10), fill='x', expand=True)
        ttk.Button(custom_frame, text="↻ Rotate", style='Accent.TButton', command=lambda: self.rotate_pdf(None)).pack(side='left', fill='x', expand=True)
        
    def _setup_watermark_tab(self):
        self.watermark_filepath = tk.StringVar()
        self.watermark_text = tk.StringVar(value="CONFIDENTIAL")
        self.watermark_opacity = tk.IntVar(value=50)
        ttk.Label(self.watermark_tab, text="Selected PDF:", font=('Arial',10,'bold')).pack(pady=(0,5), anchor='w')
        ttk.Entry(self.watermark_tab, textvariable=self.watermark_filepath, state='readonly', width=50).pack(fill='x', pady=(0,10))
        ttk.Button(self.watermark_tab, text="📂 Select PDF for Watermark", command=self.select_watermark_file).pack(fill='x', pady=5)
        ttk.Label(self.watermark_tab, text="Watermark Text:", font=('Arial',10,'bold')).pack(pady=(15,5), anchor='w')
        ttk.Entry(self.watermark_tab, textvariable=self.watermark_text, width=50).pack(fill='x', pady=5)
        ttk.Label(self.watermark_tab, text="Opacity (0-100):", font=('Arial',10,'bold')).pack(pady=(10,5), anchor='w')
        opacity_frame = ttk.Frame(self.watermark_tab)
        opacity_frame.pack(fill='x', pady=5)
        ttk.Scale(opacity_frame, from_=0, to=100, variable=self.watermark_opacity, orient='horizontal').pack(side='left', fill='x', expand=True, padx=(0,10))
        ttk.Label(opacity_frame, textvariable=self.watermark_opacity, width=3).pack(side='left')
        ttk.Button(self.watermark_tab, text="💧 Add Watermark", style='Accent.TButton', command=self.add_watermark).pack(fill='x', pady=(20, 0))
        

    def _setup_stamp_tab(self):
        self.stamp_filepath = tk.StringVar()
        self.stamp_text = tk.StringVar(value="APPROVED")
        self.stamp_color = tk.StringVar(value="#FF0000")

        self.stamp_font = tk.StringVar()
        self.stamp_font.set("Helvetica")

        self.stamp_font.trace_add("write", self.update_font_preview)
        self.stamp_color.trace_add("write", self.update_font_preview)

        ttk.Label(self.stamp_tab, text="Selected PDF:", font=('Arial',10,'bold')).pack(pady=(0,5), anchor='w')
        ttk.Entry(self.stamp_tab, textvariable=self.stamp_filepath, state='readonly', width=50).pack(fill='x', pady=(0,10))
        ttk.Button(self.stamp_tab, text="📂 Select PDF for Stamp", command=self.select_stamp_file).pack(fill='x', pady=5)

        ttk.Label(self.stamp_tab, text="Stamp Text:", font=('Arial',10,'bold')).pack(pady=(15,5), anchor='w')
        ttk.Entry(self.stamp_tab, textvariable=self.stamp_text, width=50).pack(fill='x', pady=5)

        font_frame = ttk.Frame(self.stamp_tab)
        font_frame.pack(fill='x', pady=(10,5), anchor='w')

        ttk.Label(font_frame, text="Stamp Font:", font=('Arial',10,'bold')).pack(side='left', padx=(0, 10))

        ttk.Combobox(
            font_frame,
            textvariable=self.stamp_font,
            values=self.BASE14,
            state='readonly',
            width=20
        ).pack(side='left', fill='x', expand=True, padx=(0, 10))

        self.font_preview_label = tk.Label(
            font_frame,
            text="Preview",
            fg=self.colors['PRIMARY'],
            bg=self.colors['LIGHT_BG']
        )
        self.font_preview_label.pack(side='left', padx=(5, 0), anchor='w')

        ttk.Label(self.stamp_tab, text="Stamp Color:", font=('Arial',10,'bold')).pack(pady=(10,5), anchor='w')
        color_frame = ttk.Frame(self.stamp_tab)
        color_frame.pack(fill='x', pady=5, anchor='w')

        ttk.Button(color_frame, text="Choose Color", command=self.choose_stamp_color).pack(side='left', padx=(0, 10))
        ttk.Entry(color_frame, textvariable=self.stamp_color, width=10, state='readonly').pack(side='left')

        self.color_preview_label = tk.Label(color_frame, text=" ", width=4, height=1, relief="sunken", borderwidth=1, bg=self.stamp_color.get())
        self.color_preview_label.pack(side='left', padx=(10, 0))

        ttk.Label(self.stamp_tab, text="Position:", font=('Arial',10,'bold')).pack(pady=(10,5), anchor='w')
        pos_frame = ttk.Frame(self.stamp_tab)
        pos_frame.pack(fill='x', pady=5)

        ttk.Button(pos_frame, text="Top-Left", command=lambda: self.add_stamp("top-left")).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(pos_frame, text="Center", command=lambda: self.add_stamp("center")).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(pos_frame, text="Top-Right", command=lambda: self.add_stamp("top-right")).pack(side='left', padx=5, fill='x', expand=True)
        
        self.update_font_preview()

    def update_font_preview(self, *args):
        font_name = self.stamp_font.get()
        preview_text = "Preview"
        
        weight = 'normal'
        slant = 'roman'
        
        if 'Bold' in font_name:
            weight = 'bold'
        
        if 'Oblique' in font_name or 'Italic' in font_name:
            slant = 'italic'
            
        base_font_family = font_name
        
        suffixes = ['-BoldItalic', '-Bold', 'Bold', '-Italic', 'Italic', '-Oblique', 'Oblique']
        for suffix in suffixes:
            base_font_family = base_font_family.replace(suffix, '')
            
        base_font_family = base_font_family.rstrip('-')
        
        if 'Times' in base_font_family:
            base_font = 'Times New Roman' 
        elif 'Helvetica' in base_font_family:
            base_font = 'Arial'
        elif 'Courier' in base_font_family:
            base_font = 'Courier New'
        elif 'Symbol' in base_font_family:
            base_font = 'Symbol'
        elif 'ZapfDingbats' in base_font_family:
            base_font = 'ZapfDingbats'
        else:
            base_font = 'Arial' 
        
        stamp_color_hex = self.stamp_color.get()

        try:

            self.current_preview_font = tkFont.Font(
                family=base_font, 
                size=16, 
                weight=weight, 
                slant=slant
            )
        except tk.TclError:
            self.current_preview_font = tkFont.Font(family="Arial", size=16)

        self.font_preview_label.config(
            font=self.current_preview_font,
            text=preview_text,
            fg=stamp_color_hex, 
            bg=self.colors['LIGHT_BG']
        )
        
        self.color_preview_label.config(background=stamp_color_hex)
        
    def select_encrypt_file(self):

        full_file_path = filedialog.askopenfilename(
            title="Select PDF for Encryption/Decryption",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if full_file_path:

            self._encrypt_file_full_path = full_file_path

            file_name = os.path.basename(full_file_path)
            self.encrypt_filepath.set(file_name)


    def add_merge_file(self):

        filepaths = filedialog.askopenfilenames(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")], 
            title="Select PDF files to merge"
        )
        
        if filepaths:
            for filepath in filepaths:

                self.merge_file_paths.append(filepath)
            
            self._update_listbox_display()
            
            if self.merge_file_paths:
                last_index = len(self.merge_file_paths) - 1
                self.merge_listbox.selection_clear(0, tk.END)
                self.merge_listbox.selection_set(last_index)
                self.merge_listbox.see(last_index) 

    def remove_merge_file(self):

        try:

            selected_indices = self.merge_listbox.curselection()
            
            if not selected_indices:
                messagebox.showwarning("Warning", "Please select a file to remove.")
                return
            

            index_to_remove = selected_indices[0]
            

            if 0 <= index_to_remove < len(self.merge_file_paths):
                self.merge_file_paths.pop(index_to_remove)
            

            self._update_listbox_display()

        except Exception as e:

            print(f"Error during file removal: {e}")
            messagebox.showerror("Error", "An error occurred while removing the file.")

    def select_split_file(self):
        filepath = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Select PDF file to split")
        if filepath:
            self.split_filepath.set(filepath)

    def select_extract_file(self):
        filepath = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Select PDF file")
        if filepath:
            self.extract_filepath.set(filepath)

    def merge_pdfs(self):
        if not self.merge_file_paths or len(self.merge_file_paths) < 2:
            messagebox.showwarning("Warning", "Please add at least two PDF files to merge.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Merged PDF as"
        )

        if not output_path:
            return

        try:
            output_pdf = fitz.open()

            for pdf_file in self.merge_file_paths:

                if not os.path.exists(pdf_file):
                    messagebox.showerror("Error", f"File not found:\n{pdf_file}")
                    output_pdf.close()
                    return

                doc = fitz.open(pdf_file)
                output_pdf.insert_pdf(doc)
                doc.close()

            output_pdf.save(output_path)
            output_pdf.close()

            messagebox.showinfo(
                "Success",
                f"Successfully merged {len(self.merge_file_paths)} files into:\n{output_path}"
            )

            # Clear UI after merge
            self.merge_file_paths.clear()
            self._update_listbox_display()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during merging:\n{e}")


    def split_pdf_all(self):
        filepath = self.split_filepath.get()
        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        try:
            output_dir = filedialog.askdirectory(title="Select Folder to Save Split Pages")
            if not output_dir:
                return

            doc = fitz.open(filepath)
            num_pages = len(doc)
            base_name = os.path.splitext(os.path.basename(filepath))[0]

            for i in range(num_pages):
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=i, to_page=i)
                output_filename = os.path.join(output_dir, f"{base_name}_page_{i+1}.pdf")
                new_doc.save(output_filename)
                new_doc.close()

            doc.close()
            messagebox.showinfo("Success", f"PDF successfully split into {num_pages} individual files in:\n{output_dir}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during splitting: {e}")

    def split_pdf_range(self):
        filepath = self.split_filepath.get()
        range_str = self.split_range.get().strip()

        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return
        if not range_str:
            messagebox.showwarning("Warning", "Please enter a page range (e.g., 5-10).")
            return

        try:
            start_page, end_page = map(int, range_str.split('-'))
            start_idx = start_page - 1
            end_idx = end_page - 1

            doc = fitz.open(filepath)
            num_pages = len(doc)

            if start_idx < 0 or end_idx >= num_pages or start_idx > end_idx:
                messagebox.showerror("Error", f"Invalid range. The PDF has {num_pages} pages. Ensure Start <= End.")
                doc.close()
                return

            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title=f"Save Pages {start_page} to {end_page} as")

            if output_path:
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=start_idx, to_page=end_idx)
                new_doc.save(output_path)
                new_doc.close()
                messagebox.showinfo("Success", f"Pages {start_page} to {end_page} successfully saved as:\n{output_path}")

            doc.close()

        except ValueError:
            messagebox.showerror("Error", "Invalid range format. Please use 'StartPage-EndPage' (e.g., 5-10).")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during range splitting: {e}")

    def extract_text(self):
        filepath = self.extract_filepath.get()
        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Save Extracted Text as")

        if output_path:
            try:
                doc = fitz.open(filepath)
                text_content = ""
                for page in doc:
                    text_content += page.get_text() + "\n\n-- PAGE BREAK --\n\n"
                doc.close()

                footer_watermark = (
                    "\n\n\n--- PDF TOOLKIT WATERMARK ---\n"
                    f"Generated by PDF Toolkit v{SOFTWARE_VERSION} - © {AUTHOR_NAME}\n"
                    f"Please support the author by purchasing this software. Unauthorized distribution is prohibited. Tool ID: {BUILD_ID}"
                )
                text_content += footer_watermark

                with open(output_path, "w", encoding="utf-8") as outfile:
                    outfile.write(text_content)
                messagebox.showinfo("Success", f"Text successfully extracted and saved to:\n{output_path}")

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred during text extraction: {e}")

    def extract_images(self):
        filepath = self.extract_filepath.get()
        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        output_dir = filedialog.askdirectory(title="Select Folder to Save Extracted Images")
        if not output_dir:
            return

        try:
            doc = fitz.open(filepath)
            image_count = 0

            for i in range(len(doc)):
                page = doc[i]
                for img_index, img_info in enumerate(page.get_images(full=True)):
                    xref = img_info[0]
                    base_image = doc.extract_image(xref)
                    if not base_image:
                        continue

                    image_bytes = base_image["image"]
                    ext = base_image["ext"]

                    base_name = os.path.splitext(os.path.basename(filepath))[0]
                    image_filename = os.path.join(output_dir, f"{base_name}_p{i+1}_img{img_index+1}.{ext}")

                    with open(image_filename, "wb") as f:
                        f.write(image_bytes)
                    image_count += 1

            doc.close()
            messagebox.showinfo("Success", f"Successfully extracted {image_count} images to:\n{output_dir}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during image extraction: {e}")

    def rotate_pdf(self, angle):
        filepath = getattr(self, "_rotate_file_full_path", None)
        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return
        if angle is None:
            try:
                angle = int(self.rotate_angle.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid angle.")
                return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files","*.pdf")], title="Save Rotated PDF as")
        if not output_path:
            return

        try:
            doc = fitz.open(filepath)
            for page in doc:
                page.set_rotation((page.rotation + angle) % 360)
            doc.save(output_path)
            doc.close()
            messagebox.showinfo("Success", f"PDF successfully rotated by {angle}° and saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during rotation: {e}")

    def create_watermark_pdf(text="CONFIDENTIAL", opacity=0.3, angle=45):
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        c.saveState()
        c.setFillColor(Color(0, 0, 0, alpha=opacity))
        c.translate(300, 400)
        c.rotate(angle)
        c.setFont("Helvetica-Bold", 60)
        c.drawCentredString(0, 0, text)
        c.restoreState()
        c.save()
        packet.seek(0)
        return PdfReader(packet)

    def add_watermark(self):
        if not hasattr(self, '_watermark_file_full_path') or not self._watermark_file_full_path:
            messagebox.showwarning("Warning", "Please select a PDF first.")
            return
            
        input_pdf = getattr(self, "_watermark_file_full_path", None)
        if not input_pdf:
            messagebox.showwarning("Warning", "Please select a PDF first.")
            return

        if not os.path.exists(input_pdf):
            messagebox.showerror("Error", f"The selected file was not found:\n{input_pdf}")
            return


        output_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Watermarked PDF As"
        )
        if not output_pdf:
            return

        watermark_text = self.watermark_text.get().strip()
        if not watermark_text:
            messagebox.showwarning("Warning", "Please enter watermark text.")
            return

        opacity = self.watermark_opacity.get() / 100.0
        angle = 45

        try:
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            for page_num, page in enumerate(reader.pages):
                packet = io.BytesIO()

                width = float(page.mediabox.width)
                height = float(page.mediabox.height)

                c = canvas.Canvas(packet, pagesize=(width, height))
                c.setFillColor(Color(0.5, 0.5, 0.5, alpha=opacity))

                c.saveState()
                c.translate(width/2, height/2)
                c.rotate(angle)
                c.setFont("Helvetica", 60)
                c.drawCentredString(0, 0, watermark_text)
                c.restoreState()
                c.save()

                packet.seek(0)
                watermark_pdf = PdfReader(packet)
                page.merge_page(watermark_pdf.pages[0])
                writer.add_page(page)

            with open(output_pdf, "wb") as f_out:
                writer.write(f_out)

            messagebox.showinfo("Success", f"Watermark added successfully!\nSaved to: {output_pdf}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to add watermark:\n{str(e)}")

            
    def choose_stamp_color(self):
        from tkinter import colorchooser

        color_code = colorchooser.askcolor(
            title="Choose Stamp Color",
            initialcolor=self.stamp_color.get()
        )
        if color_code and color_code[1]:
            self.stamp_color.set(color_code[1])

    def is_valid_embed_font(self, path):
        if not path or not os.path.exists(path):
            return False
        if not path.lower().endswith((".ttf", ".otf")):
            return False
        try:
            with open(path, "rb") as f:
                f.seek(0, os.SEEK_END)
                if f.tell() < 2000:
                    return False
        except:
            return False
        return True

    def add_stamp(self, position):
        filepath = getattr(self, "_stamp_file_full_path", None)
        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return
        

        if not os.path.exists(filepath):
            messagebox.showerror("Error", f"The selected file was not found:\n{filepath}")
            return


        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Stamped PDF as"
        )
        if not output_path:
            return

        try:
            doc = fitz.open(filepath)
            text = self.stamp_text.get()
            color = self.stamp_color.get()
            font_name = self.stamp_font.get()
            

            font_path = getattr(self, '_custom_stamp_font_path', None)


            BASE14 = [
                "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
                "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
                "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique",
                "Symbol", "ZapfDingbats"
            ]
            py_font = None

            if font_name in BASE14:
                py_font = font_name
            else:
                if font_path:

                    if not os.path.exists(font_path):
                        messagebox.showwarning("Warning", f"Custom font not found: {os.path.basename(font_path)}. Falling back to Helvetica.")
                        py_font = "Helvetica"
                    else:
                        alias = f"embedded_{os.path.basename(font_path)}"
                        doc.insert_font(fontfile=font_path, fontname=alias)
                        py_font = alias
                else:
                    py_font = "Helvetica"

            rgb = tuple(int(color[i:i+2], 16) / 255 for i in (1, 3, 5))

            for page in doc:
                rect = page.rect
                x, y = rect.width / 2, rect.height / 2
                if position == "top-left":
                    x, y = 50, 50
                elif position == "top-right":
                    x, y = rect.width - 50, 50

                page.insert_text(
                    (x, y),
                    text,
                    fontsize=36,
                    color=rgb,
                    fontname=py_font,
                    rotate=0,
                    overlay=True
                )

            doc.save(output_path)
            doc.close()
            messagebox.showinfo("Success", f"Stamp added and saved to:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during stamping:\n{e}")

    def toggle_password_visibility(self, *args):

        if hasattr(self, 'encrypt_password_entry'):
            show_char = '' if self.show_password_var.get() else '*'
            self.encrypt_password_entry.config(show=show_char)
            self.encrypt_confirm_password_entry.config(show=show_char)
            
        if hasattr(self, 'decrypt_password_entry'):
            show_char = '' if self.show_password_var.get() else '*'
            self.decrypt_password_entry.config(show=show_char)


    def encrypt_pdf(self):
        filepath = getattr(self, "_encrypt_file_full_path", None)
        password = self.encrypt_password.get()
        confirm_password = self.encrypt_confirm_password.get()

        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        if not password:
            messagebox.showwarning("Warning", "Please enter a password.")
            return
        
        if password != confirm_password:
            messagebox.showerror("Error", "Passwords do not match. Please re-enter.")
            return

        if len(password) < 4:
            messagebox.showwarning("Warning", "Password should be at least 4 characters long.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Encrypted PDF as")
        
        if not output_path:
            return

        try:
            reader = PdfReader(filepath)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            writer.encrypt(password, password)

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            self.encrypt_password.set("")
            self.encrypt_confirm_password.set("")
            self.decrypt_password.set("")
            self.show_password_var.set(False) 
            
            messagebox.showinfo("Success", f"PDF successfully encrypted and saved to:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during encryption: {e}")
            
    def decrypt_pdf(self):
        filepath = getattr(self, "_encrypt_file_full_path", None)
        password = self.decrypt_password.get() 

        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        if not password:
            messagebox.showwarning("Warning", "Please enter the decryption password.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Decrypted PDF as")
        
        if not output_path:
            return

        try:
            reader = PdfReader(filepath)
            
            if reader.is_encrypted:
                try:
                    reader.decrypt(password)
                except Exception:
                    messagebox.showerror("Error", "Incorrect password for decryption.")
                    return
            else:

                if not messagebox.askyesno("Not Encrypted", "The selected PDF is not encrypted. Do you want to save an unencrypted copy?"):
                    return
            
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            with open(output_path, "wb") as output_file:
                writer.write(output_file)
            
            self.decrypt_password.set("")
            self.encrypt_password.set("")
            self.encrypt_confirm_password.set("")
            self.show_password_var.set(False) 
            
            messagebox.showinfo("Success", f"PDF successfully decrypted/copied and saved to:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred during decryption: {e}")


    def select_rotate_file(self):

        full_file_path = filedialog.askopenfilename(
            title="Select PDF for Rotation",
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if full_file_path:

            self._rotate_file_full_path = full_file_path
            

            file_name = os.path.basename(full_file_path)
            self.rotate_filepath.set(file_name)
            
    def select_watermark_file(self):

        full_file_path = filedialog.askopenfilename(
            title="Select PDF for Watermarking",
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if full_file_path:

            self._watermark_file_full_path = full_file_path
            
            file_name = os.path.basename(full_file_path)
            self.watermark_filepath.set(file_name)

    def select_stamp_file(self):

        full_file_path = filedialog.askopenfilename(
            title="Select PDF for Stamping",
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if full_file_path:

            self._stamp_file_full_path = full_file_path
            
            file_name = os.path.basename(full_file_path)
            self.stamp_filepath.set(file_name)

    def _setup_metadata_tab(self):

        ttk.Label(self.metadata_tab, text="Selected PDF:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor='w')
        ttk.Entry(self.metadata_tab, textvariable=self.metadata_filepath, state='readonly', width=50).pack(fill='x', pady=(0, 10))
        ttk.Button(self.metadata_tab, text="📂 Select PDF & Load Metadata", command=self.select_metadata_file).pack(fill='x', pady=5)
        
        fields_frame = ttk.Frame(self.metadata_tab)
        fields_frame.pack(fill='x', pady=(15, 0))
        
        def add_metadata_field(parent, label_text, var):
            row_frame = ttk.Frame(parent)
            row_frame.pack(fill='x', pady=5)
            ttk.Label(row_frame, text=label_text, width=15, anchor='w').pack(side='left')
            ttk.Entry(row_frame, textvariable=var).pack(side='left', fill='x', expand=True)

        add_metadata_field(fields_frame, "Title:", self.metadata_title)
        add_metadata_field(fields_frame, "Author:", self.metadata_author)
        add_metadata_field(fields_frame, "Subject:", self.metadata_subject)
        add_metadata_field(fields_frame, "Creator:", self.metadata_creator)

        ttk.Button(self.metadata_tab, text="💾 Save New Metadata", style='Accent.TButton', command=self.save_metadata).pack(fill='x', pady=(20, 0))
        
        ttk.Label(
            self.metadata_tab, 
            text="Note: Saving metadata will create a new PDF file.",
            foreground="#777777",
            font=('Arial', 9, 'italic')
        ).pack(pady=(10, 0), anchor='w')

    def select_metadata_file(self):
        full_file_path = filedialog.askopenfilename(
            title="Select PDF to Edit Metadata",
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if full_file_path:
            self._metadata_file_full_path = full_file_path
            file_name = os.path.basename(full_file_path)
            self.metadata_filepath.set(file_name)
            self.load_metadata(full_file_path)

    def load_metadata(self, filepath):

        try:
            reader = PdfReader(filepath)
            meta = reader.metadata
            
            self.metadata_title.set(meta.get('/Title', ''))
            self.metadata_author.set(meta.get('/Author', ''))
            self.metadata_subject.set(meta.get('/Subject', ''))
            self.metadata_creator.set(meta.get('/Creator', ''))
            
            messagebox.showinfo("Metadata Loaded", "Metadata loaded successfully from the PDF.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load metadata: {e}")
            self.metadata_title.set('')
            self.metadata_author.set('')
            self.metadata_subject.set('')
            self.metadata_creator.set('')

    def save_metadata(self):
        filepath = getattr(self, "_metadata_file_full_path", None)
        if not filepath:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF with New Metadata as"
        )
        
        if not output_path:
            return

        try:
            reader = PdfReader(filepath)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            new_meta = {
                '/Title': self.metadata_title.get(),
                '/Author': self.metadata_author.get(),
                '/Subject': self.metadata_subject.get(),
                '/Creator': self.metadata_creator.get(),
            }
            
            writer.add_metadata(new_meta)

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            messagebox.showinfo("Success", f"Metadata updated and saved to:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the metadata: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolkitApp(root)
    root.mainloop()
