import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image
from pillow_heif import register_heif_opener
import threading
import io
import time
from docx2pdf import convert as docx_to_pdf
import pandas as pd

register_heif_opener()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

CONVERSION_CONFIG = {
    "Kompresi": {
        "formats": ["JPG"],
        "output_map": {
            "JPG": ["JPG (Kualitas Dikompresi)"]
        }
    },
    "Gambar": {
        "formats": ["HEIC", "JPG", "PNG", "WEBP"],
        "output_map": {
            "HEIC": ["JPG", "PNG"],
            "JPG": ["PNG", "WEBP", "ICO"],
            "PNG": ["JPG", "WEBP", "ICO"],
            "WEBP": ["JPG", "PNG"]
        }
    },
    "Dokumen": {
        "formats": ["DOCX", "XLSX", "CSV"],
        "output_map": {
            "DOCX": ["PDF"],
            "XLSX": ["CSV"],
            "CSV": ["XLSX"]
        }
    }
}

def convert_file(input_path, output_path, from_format, to_format, category, target_size_kb=None):
    """Fungsi dispatcher untuk memanggil metode konversi/kompresi yang benar."""
    try:
        if category == "Kompresi":
            if from_format == "JPG":
                target_bytes = int(target_size_kb) * 1024
                with Image.open(input_path) as img:
                    if img.mode in ('RGBA', 'P'):
                        img = img.convert('RGB')
                    
                    low = 1
                    high = 95
                    best_quality = -1
                    
                    while low <= high:
                        mid = (low + high) // 2
                        buffer = io.BytesIO()
                        img.save(buffer, "JPEG", quality=mid, optimize=True)
                        size = buffer.tell()
                        
                        if size <= target_bytes:
                            best_quality = mid
                            low = mid + 1
                        else:
                            high = mid - 1

                    if best_quality != -1:
                        buffer = io.BytesIO()
                        img.save(buffer, "JPEG", quality=best_quality, optimize=True)
                        with open(output_path, 'wb') as f:
                            f.write(buffer.getvalue())
                    else:
                        img.save(output_path, "JPEG", quality=1, optimize=True)

        elif category == "Gambar":
            with Image.open(input_path) as img:
                if to_format == "JPG":
                    if img.mode in ('RGBA', 'P'):
                        img = img.convert('RGB')
                    img.save(output_path, "JPEG", quality=95)
                elif to_format == "ICO":
                    icon_sizes = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
                    img.save(output_path, format='ICO', sizes=icon_sizes)
                else:
                    img.save(output_path, format=to_format.upper())
        
        elif category == "Dokumen":
            if from_format == "DOCX" and to_format == "PDF":
                docx_to_pdf(input_path, output_path)
            elif from_format == "XLSX" and to_format == "CSV":
                pd.read_excel(input_path).to_csv(output_path, index=False)
            elif from_format == "CSV" and to_format == "XLSX":
                pd.read_csv(input_path).to_excel(output_path, index=False)

        return True

    except Exception:
        return False

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter & Compressor Pro")
        
        try:
            self.root.iconbitmap(resource_path('iconCONVERT.ico'))
        except tk.TclError:
            pass

        self.root.geometry("780x900")
        self.root.resizable(False, False)
        self.root.minsize(700, 600)

        self.list_of_files = []
        
        self.output_folder_path = tk.StringVar()
        self.files_label_var = tk.StringVar(value="Belum ada file yang dipilih.")
        self.status_var = tk.StringVar(value="🎉 Selamat datang! Silakan pilih jenis operasi untuk memulai.")
        self.category_var = tk.StringVar()
        self.from_format_var = tk.StringVar()
        self.to_format_var = tk.StringVar()
        self.target_size_var = tk.StringVar(value="1024")
        self.progress_text_var = tk.StringVar(value="")
        
        self.create_widgets()
        self.update_format_options()

    def create_widgets(self):
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=BOTH, expand=True)
        
        self.canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        main_frame = ttk.Frame(self.scrollable_frame, padding="30")
        main_frame.pack(fill=BOTH, expand=True)

        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=X, pady=(0, 30))
        
        title_label = ttk.Label(
            header_frame, 
            text="🚀 File Converter & Compressor Pro", 
            font=("Segoe UI", 24, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame, 
            text="Konversi dan kompresi file dengan mudah dan cepat",
            font=("Segoe UI", 12),
            foreground="#7f8c8d"
        )
        subtitle_label.pack(pady=(5, 0))
        
        options_card = ttk.Labelframe(
            main_frame, 
            text="🔧 Pengaturan Konversi", 
            padding=25,
            bootstyle="primary"
        )
        options_card.pack(fill=X, pady=(0, 20))
        options_card.grid_columnconfigure(1, weight=1)

        cat_frame = ttk.Frame(options_card)
        cat_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        cat_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(cat_frame, text="📂 Kategori:", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.category_menu = ttk.Combobox(
            cat_frame, 
            textvariable=self.category_var, 
            state="readonly", 
            values=list(CONVERSION_CONFIG.keys()),
            font=("Segoe UI", 10),
            width=25
        )
        self.category_menu.grid(row=0, column=1, sticky="ew")

        from_frame = ttk.Frame(options_card)
        from_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        from_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(from_frame, text="📥 Dari Format:", font=("Segoe UI", 11)).grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.from_format_menu = ttk.Combobox(
            from_frame, 
            textvariable=self.from_format_var, 
            state="disabled",
            font=("Segoe UI", 10),
            width=25
        )
        self.from_format_menu.grid(row=0, column=1, sticky="ew")

        to_frame = ttk.Frame(options_card)
        to_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        to_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(to_frame, text="📤 Ke Format:", font=("Segoe UI", 11)).grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.to_format_menu = ttk.Combobox(
            to_frame, 
            textvariable=self.to_format_var, 
            state="disabled",
            font=("Segoe UI", 10),
            width=25
        )
        self.to_format_menu.grid(row=0, column=1, sticky="ew")
        
        self.target_size_frame = ttk.Frame(options_card)
        self.target_size_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        self.target_size_frame.grid_columnconfigure(1, weight=1)
        
        self.target_size_label = ttk.Label(self.target_size_frame, text="🎯 Target Ukuran (KB):", font=("Segoe UI", 11))
        self.target_size_label.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.target_size_entry = ttk.Entry(
            self.target_size_frame, 
            textvariable=self.target_size_var,
            font=("Segoe UI", 10),
            width=25
        )
        self.target_size_entry.grid(row=0, column=1, sticky="ew")

        self.category_var.trace("w", self.update_format_options)
        self.from_format_var.trace("w", self.update_output_options)
        self.to_format_var.trace("w", lambda *_: self.check_and_enable_button())

        io_main_frame = ttk.Frame(main_frame)
        io_main_frame.pack(fill=X, pady=(0, 20))
        
        files_card = ttk.Labelframe(
            io_main_frame, 
            text="📁 File Sumber", 
            padding=20,
            bootstyle="info"
        )
        files_card.pack(fill=BOTH, expand=True, side=LEFT, padx=(0, 10))
        
        select_files_button = ttk.Button(
            files_card, 
            text="🔍 Pilih File", 
            command=self.select_files, 
            bootstyle=(PRIMARY, OUTLINE),
            width=15
        )
        select_files_button.pack(pady=(0, 15))
        
        self.files_display_frame = ttk.Frame(files_card)
        self.files_display_frame.pack(fill=BOTH, expand=True)
        
        self.files_label = ttk.Label(
            self.files_display_frame, 
            textvariable=self.files_label_var, 
            wraplength=300, 
            justify=CENTER,
            font=("Segoe UI", 10),
            foreground="#34495e"
        )
        self.files_label.pack(pady=10)

        output_card = ttk.Labelframe(
            io_main_frame, 
            text="📂 Folder Output", 
            padding=20,
            bootstyle="success"
        )
        output_card.pack(fill=BOTH, expand=True, side=RIGHT, padx=(10, 0))
        
        select_folder_button = ttk.Button(
            output_card, 
            text="📁 Pilih Folder", 
            command=self.select_output_folder, 
            bootstyle=(SUCCESS, OUTLINE),
            width=15
        )
        select_folder_button.pack(pady=(0, 15))
        
        self.output_folder_entry = ttk.Entry(
            output_card, 
            textvariable=self.output_folder_path, 
            state="readonly", 
            font=("Segoe UI", 9),
            justify=CENTER
        )
        self.output_folder_entry.pack(fill=X, pady=10)

        process_card = ttk.Labelframe(
            main_frame, 
            text="⚡ Proses Konversi", 
            padding=25,
            bootstyle="warning"
        )
        process_card.pack(fill=X, pady=(0, 20))

        self.convert_button = ttk.Button(
            process_card, 
            text="🚀 Mulai Proses Konversi", 
            command=self.start_conversion_thread, 
            bootstyle=(SUCCESS, "outline-toolbutton"),
            state="disabled",
            width=30
        )
        self.convert_button.pack(pady=(0, 20))

        progress_frame = ttk.Frame(process_card)
        progress_frame.pack(fill=X)
        
        self.progress_label = ttk.Label(
            progress_frame, 
            textvariable=self.progress_text_var,
            font=("Segoe UI", 10, "bold"),
            foreground="#e74c3c"
        )
        self.progress_label.pack(pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            mode='determinate', 
            bootstyle=(SUCCESS, STRIPED),
            length=400
        )
        self.progress_bar.pack(fill=X, pady=(0, 15))

        status_card = ttk.Frame(main_frame)
        status_card.pack(fill=X, pady=(0, 10))
        
        status_inner = ttk.Label(
            status_card,
            textvariable=self.status_var, 
            font=("Segoe UI", 11),
            foreground="#2c3e50",
            background="#ecf0f1",
            relief="solid",
            borderwidth=1,
            padding=(15, 10)
        )
        status_inner.pack(fill=X)

        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.root.bind("<Configure>", self._on_window_resize)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def _on_window_resize(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def update_format_options(self, *args):
        category = self.category_var.get()
        formats = CONVERSION_CONFIG.get(category, {}).get("formats", [])
        self.from_format_menu.config(values=formats, state="readonly" if formats else "disabled")
        self.from_format_var.set("")
        self.to_format_var.set("")
        self.check_and_enable_button()
        self.update_output_options()
        
        if category:
            self.status_var.set(f"✅ Kategori '{category}' dipilih. Silakan pilih format input dan output.")
    
    def update_output_options(self, *args):
        category = self.category_var.get()
        from_format = self.from_format_var.get()
        output_formats = CONVERSION_CONFIG.get(category, {}).get("output_map", {}).get(from_format, [])
        self.to_format_menu.config(values=output_formats, state="readonly" if output_formats else "disabled")
        self.to_format_var.set("")

        if category == "Kompresi" and from_format == "JPG":
            self.target_size_frame.grid()
        else:
            self.target_size_frame.grid_remove()
            
        self.check_and_enable_button()

    def select_files(self):
        from_format = self.from_format_var.get()
        if not from_format:
            messagebox.showwarning("⚠️ Peringatan", "Harap pilih format input terlebih dahulu.")
            return
            
        file_extension = f"*.{from_format.lower()}"
        file_type_desc = f"{from_format} Files"
        
        filepaths = filedialog.askopenfilenames(
            title=f"Pilih File {from_format}",
            filetypes=[(file_type_desc, file_extension), ("All Files", "*.*")]
        )
        
        if filepaths:
            self.list_of_files = list(filepaths)
            self.files_label_var.set(f"✅ {len(self.list_of_files)} file dipilih")
            self.status_var.set(f"📂 {len(self.list_of_files)} file siap untuk diproses.")
            self.check_and_enable_button()

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="Pilih Folder Output")
        if folder_path:
            self.output_folder_path.set(folder_path)
            self.status_var.set("📁 Folder output telah dipilih.")
            self.check_and_enable_button()

    def start_conversion_thread(self):
        """Mulai proses di thread terpisah agar GUI tidak macet."""
        self.convert_button.config(state="disabled", text="⏳ Sedang Memproses...")
        self.progress_bar['value'] = 0
        self.progress_text_var.set("🔄 Mempersiapkan proses konversi...")
        
        conversion_thread = threading.Thread(target=self.run_conversion, daemon=True)
        conversion_thread.start()

    def run_conversion(self):
        from_format = self.from_format_var.get()
        to_format = self.to_format_var.get()
        category = self.category_var.get()
        
        if not all([self.list_of_files, self.output_folder_path.get(), from_format, to_format]):
            messagebox.showerror("❌ Error", "Harap lengkapi semua pilihan sebelum memulai proses.")
            self.reset_button_state()
            return
        
        target_size_kb = None
        if category == "Kompresi" and from_format == "JPG":
            try:
                target_size_kb = int(self.target_size_var.get())
                if target_size_kb <= 0:
                    messagebox.showerror("❌ Error", "Target ukuran harus angka positif.")
                    self.reset_button_state()
                    return
            except ValueError:
                messagebox.showerror("❌ Error", "Target ukuran harus berupa angka.")
                self.reset_button_state()
                return

        total_files = len(self.list_of_files)
        self.progress_bar['maximum'] = total_files
        success_count = 0
        fail_count = 0

        start_time = time.time()

        for i, input_path in enumerate(self.list_of_files):
            filename = os.path.basename(input_path)
            
            if category == "Kompresi":
                base, ext = os.path.splitext(filename)
                output_filename = f"{base}_terkompresi{ext}"
            else:
                output_filename = os.path.splitext(filename)[0] + f".{to_format.lower()}"

            output_path = os.path.join(self.output_folder_path.get(), output_filename)
            
            self.progress_text_var.set(f"🔄 Memproses ({i+1}/{total_files}): {filename}")
            self.status_var.set(f"⚡ Sedang memproses file ke-{i+1} dari {total_files}")
            
            if convert_file(input_path, output_path, from_format, to_format, category, target_size_kb=target_size_kb):
                success_count += 1
            else:
                fail_count += 1
            
            self.progress_bar['value'] = i + 1
            self.root.update_idletasks()
            
            time.sleep(0.1)

        end_time = time.time()
        duration = round(end_time - start_time, 2)
        
        self.progress_text_var.set("✅ Proses konversi selesai!")
        self.status_var.set(f"🎉 Proses selesai dalam {duration} detik.")
        
        if fail_count == 0:
            icon = "🎉"
            title = "Berhasil!"
            message = f"{icon} Semua file berhasil diproses!\n\n✅ Berhasil: {success_count}\n⏱️ Waktu: {duration} detik"
        else:
            icon = "⚠️"
            title = "Proses Selesai"
            message = f"{icon} Proses Selesai!\n\n✅ Berhasil: {success_count}\n❌ Gagal: {fail_count}\n⏱️ Waktu: {duration} detik"
            
        messagebox.showinfo(title, message)
        
        self.reset_ui()

    def reset_button_state(self):
        self.convert_button.config(state="normal", text="🚀 Mulai Proses Konversi")
        self.progress_text_var.set("")

    def check_and_enable_button(self):
        if self.list_of_files and self.output_folder_path.get() and self.to_format_var.get():
            self.convert_button.config(state="normal")
            self.status_var.set("🚀 Siap untuk memulai konversi!")
        else:
            self.convert_button.config(state="disabled")
            
    def reset_ui(self):
        self.progress_bar['value'] = 0
        self.progress_text_var.set("")
        self.files_label_var.set("Belum ada file yang dipilih.")
        self.output_folder_path.set("")
        self.list_of_files = []
        self.status_var.set("🎉 Selamat datang! Silakan pilih jenis operasi untuk memulai.")
        self.convert_button.config(state="disabled", text="🚀 Mulai Proses Konversi")

if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = ConverterApp(root)
    root.mainloop()
