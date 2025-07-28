"""
Main GUI application for SIKELAR
Handles user interface and interactions
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List
import sys
import os

# Add parent directory to path to import backend modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.processor import BOSDataProcessor
from backend.utils import FormatUtils

class BOSBudgetAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("SIKELAR")
        self.root.geometry("1000x700")
        self.root.configure(bg='#f0f0f0')
        
        # Initialize data processor
        self.processor = BOSDataProcessor()
        
        # For supporting active button highlight
        self.tab_buttons = {}
        self.active_tab = None

        self.setup_ui()

    def setup_ui(self):
        """Setup the main user interface"""
        self._create_header()
        self._create_upload_section()
        self._create_button_section()
        self._create_results_section()

    def _create_header(self):
        """Create header section"""
        header_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        header_frame.pack(fill='x', padx=10, pady=(10,0))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="Sistem Informasi Pengelompokan Anggaran dan Rekening", 
                            font=('Arial', 16, 'bold'), fg='white', bg='#2c3e50')
        title_label.pack(expand=True)

    def _create_upload_section(self):
        """Create upload section"""
        upload_frame = tk.Frame(self.root, bg='#ecf0f1', relief='raised', bd=2)
        upload_frame.pack(fill='x', padx=10, pady=10)
        
        upload_label = tk.Label(upload_frame, text="Upload File Excel RKAS:", 
                            font=('Arial', 12, 'bold'), bg='#ecf0f1')
        upload_label.pack(pady=10)
        
        upload_btn_frame = tk.Frame(upload_frame, bg="#ecf0f1")
        upload_btn_frame.pack(pady=10)

        common_font = ('Arial', 11, 'bold')

        self.upload_btn = tk.Button(upload_btn_frame, text="Pilih File Excel (.xlsx)", command=self.upload_excel,
                            bg='#3498db', fg='white', font=common_font,
                            padx=20, pady=5, cursor='hand2')
        self.upload_btn.pack(side='left', padx=(0, 5))

        reset_btn = tk.Button(upload_btn_frame, text="✖", command=self.reset_data,
                            bg='red', fg='white', font=common_font,
                            padx=5, pady=5, cursor='hand2', bd=1, relief='raised')
        reset_btn.pack(side='left')
        
        self.file_label = tk.Label(upload_frame, text="Belum ada file yang dipilih", 
                                font=('Arial', 10), bg='#ecf0f1', fg='#7f8c8d')
        self.file_label.pack(pady=5)

    def _create_button_section(self):
        """Create navigation button section"""
        main_button_frame = tk.Frame(self.root, bg='#f0f0f0', height=80)
        main_button_frame.pack(fill='x', padx=10, pady=10)
        main_button_frame.pack_propagate(False)
        
        self.button_canvas = tk.Canvas(main_button_frame, bg='#f0f0f0', height=60, highlightthickness=0)
        self.button_canvas.pack(side='top', fill='x', padx=5, pady=5)
        
        h_scrollbar = tk.Scrollbar(main_button_frame, orient='horizontal', command=self.button_canvas.xview)
        h_scrollbar.pack(side='bottom', fill='x')
        self.button_canvas.configure(xscrollcommand=h_scrollbar.set)
        
        self.scrollable_button_frame = tk.Frame(self.button_canvas, bg='#f0f0f0')
        self.canvas_window = self.button_canvas.create_window((0, 0), window=self.scrollable_button_frame, anchor='nw')
        
        self._create_navigation_buttons()
        self._setup_canvas_scrolling()

    def _create_navigation_buttons(self):
        """Create navigation buttons"""
        button_font = ('Arial', 11, 'bold')
        
        buttons_config = [
            ("Belanja Persediaan", self.show_belanja_persediaan, "#29b9b9", 24),
            ("Jasa", self.show_belanja_jasa, "#29b9b9", 24),
            ("Pemeliharaan", self.show_belanja_pemeliharaan, "#29b9b9", 24),
            ("Perjalanan Dinas", self.show_belanja_perjalanan, "#29b9b9", 24),
            ("Peralatan", self.show_peralatan,  "#29b9b9", 24),
            ("Aset Tetap", self.show_aset_tetap, "#29b9b9", 24),
            ("Ringkasan", self.show_ringkasan, "#29b9b9", 24)
        ]
        
        for name, func, color, width in buttons_config:
            def make_command(f=func, n=name):
                return lambda: [f(), self._highlight_active_tab(n)]

            btn = tk.Button(self.scrollable_button_frame, 
                        text=name, 
                        command=make_command(), 
                        bg=color, 
                        fg='white', 
                        font=button_font, 
                        cursor='hand2', 
                        width=width,
                        height=2,
                        relief='raised',
                        bd=2)
            btn.pack(side='left', padx=3, pady=5)
            self.tab_buttons[name] = (btn, color)

    def _highlight_active_tab(self, name):
        """Highlight the active tab button"""
        for tab_name, (btn, orig_color) in self.tab_buttons.items():
            if tab_name == name:
                btn.config(bg=FormatUtils.darken_color(orig_color))
            else:
                btn.config(bg=orig_color)
        self.active_tab = name

    def _setup_canvas_scrolling(self):
        """Setup canvas scrolling functionality"""
        self.scrollable_button_frame.update_idletasks()
        self.button_canvas.configure(scrollregion=self.button_canvas.bbox("all"))
        
        self.button_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.button_canvas.bind("<Button-4>", self._on_mousewheel)
        self.button_canvas.bind("<Button-5>", self._on_mousewheel)
        self.button_canvas.bind('<Configure>', self._on_canvas_configure)
        self.scrollable_button_frame.bind('<Configure>', self._on_frame_configure)
        self.button_canvas.focus_set()

    def _create_results_section(self):
        """Create results display section"""
        self.results_frame = tk.Frame(self.root, bg='#ffffff', relief='sunken', bd=2)
        self.results_frame.pack(fill='both', expand=True, padx=10, pady=10)

    def upload_excel(self):
        """Handle Excel file upload"""
        file_path = filedialog.askopenfilename(title="Pilih File Excel RKAS", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.processor.extract_excel_data(file_path)
                self.file_label.config(text=f"File dipilih: {file_path.split('/')[-1]}", fg='#27ae60')

                # Disable upload button after successful upload
                self.upload_btn.config(state='disabled')

                messagebox.showinfo("Berhasil", 
                    f"File Excel berhasil diproses!\n"
                    f"Total Penerimaan: {FormatUtils.format_currency(self.processor.total_penerimaan)}\n"
                    f"Ditemukan {len(self.processor.belanja_persediaan_items)} item belanja persediaan\n"
                    f"Ditemukan {len(self.processor.belanja_jasa_items)} item belanja jasa\n"
                    f"Ditemukan {len(self.processor.belanja_pemeliharaan_items)} item belanja pemeliharaan\n"
                    f"Ditemukan {len(self.processor.belanja_perjalanan_items)} item belanja perjalanan\n"
                    f"Ditemukan {len(self.processor.peralatan_items)} item peralatan dan mesin\n"
                    f"Ditemukan {len(self.processor.aset_tetap_items)} item aset tetap lainnya")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal membaca file Excel: {str(e)}")

    def reset_data(self):
        """Reset all data and UI"""
        self.processor.reset_data()

        # Clear results display
        for widget in self.results_frame.winfo_children():
            widget.destroy()

        # Clear file label
        self.file_label.config(text="Belum ada file yang dipilih")

        # Re-enable upload button
        self.upload_btn.config(state='normal')

        messagebox.showinfo("Reset", "Data berhasil dibersihkan.")

    def create_standard_table(self, title, columns):
        """Create standard table layout"""
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        self.result_title = tk.Label(self.results_frame, text=title, 
                                    font=('Arial', 16, 'bold'), bg='#ffffff')
        self.result_title.pack(pady=20)
        
        self.table_frame = tk.Frame(self.results_frame, bg='#ffffff')
        self.table_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show='headings', height=15)
        
        # Style untuk memperbesar font tabel
        style = ttk.Style()
        style.configure("Treeview", font=('Arial', 12))
        style.configure("Treeview.Heading", font=('Arial', 14, 'bold'))
        
        for col in columns:
            self.tree.heading(col, text=col)
        
        # Set column widths based on content
        if len(columns) == 4:  # Standard detail table
            self.tree.column(columns[0], width=150, anchor='center')  # Kode Rekening
            self.tree.column(columns[1], width=120, anchor='center')  # Kode Kegiatan
            self.tree.column(columns[2], width=350, anchor='w')       # Uraian
            self.tree.column(columns[3], width=150, anchor='w')       # Jumlah
        elif len(columns) == 2:  # Summary table
            self.tree.column(columns[0], width=400, anchor='w')       # Kategori
            self.tree.column(columns[1], width=200, anchor='w')       # Jumlah
        
        scrollbar = ttk.Scrollbar(self.table_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.summary_frame = tk.Frame(self.results_frame, bg='#ecf0f1', relief='raised', bd=1)
        self.summary_frame.pack(fill='x', padx=20, pady=10)

    def display_standard_results(self, items: List[Dict], title: str, total_label: str):
        """Display results in standard format"""
        columns = ('Kode Rekening', 'Kode Kegiatan', 'Uraian', 'Jumlah (Rp)')
        self.create_standard_table(title, columns)
        
        total_jumlah = 0
        if items:
            for item in items:
                formatted_jumlah = FormatUtils.format_currency(item['jumlah'])
                self.tree.insert('', 'end', values=(
                    item['kode_rekening'], 
                    item['kode_kegiatan'], 
                    item['uraian'], 
                    formatted_jumlah
                ))
                total_jumlah += item['jumlah']
        else:
            self.tree.insert('', 'end', values=('', '', 'Tidak ada data ditemukan untuk kategori ini', 'Rp 0'))
        
        # Create summary labels
        self._create_standard_summary(total_label, total_jumlah)

    def _create_standard_summary(self, total_label: str, total_jumlah: int):
        """Create standard summary section"""
        # Frame kiri untuk nama sekolah
        left_frame = tk.Frame(self.summary_frame, bg='#ecf0f1')
        left_frame.pack(side='left', fill='y', padx=(10, 20))
        
        sekolah_label = tk.Label(left_frame, text=f"SEKOLAH {self.processor.nama_sekolah}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        sekolah_label.pack(anchor='w')
        
        self.total_label = tk.Label(self.summary_frame, text=f"{total_label}: {FormatUtils.format_currency(total_jumlah)}", 
                                   font=('Arial', 14, 'bold'), bg='#ecf0f1')
        self.total_label.pack(pady=5)

    def display_belanja_jasa_results(self, items: List[Dict]):
        """Display belanja jasa results with extended summary"""
        columns = ('Kode Rekening', 'Kode Kegiatan', 'Uraian', 'Jumlah (Rp)')
        self.create_standard_table("Rincian Belanja Jasa (5.1.02.02)", columns)
        
        total_belanja_jasa = 0
        if items:
            for item in items:
                formatted_jumlah = FormatUtils.format_currency(item['jumlah'])
                self.tree.insert('', 'end', values=(
                    item['kode_rekening'], 
                    item['kode_kegiatan'], 
                    item['uraian'], 
                    formatted_jumlah
                ))
                total_belanja_jasa += item['jumlah']
        else:
            self.tree.insert('', 'end', values=('', '', 'Tidak ada data ditemukan untuk kategori ini', 'Rp 0'))
        
        # Calculate honor and actual service
        honor_items = self.processor.filter_budget_by_codes(self.processor.kategori_kode['honor'])
        total_honor = sum(item['jumlah'] for item in honor_items)
        jasa_sesungguhnya = total_belanja_jasa - total_honor
        
        # Create extended summary
        self.total_label = tk.Label(self.summary_frame, text=f"Total Belanja Jasa: {FormatUtils.format_currency(total_belanja_jasa)}", 
                                   font=('Arial', 14, 'bold'), bg='#ecf0f1')
        self.total_label.pack(pady=5)
        
        self.honor_label = tk.Label(self.summary_frame, 
                                text=f"Pembayaran Honor: {FormatUtils.format_currency(total_honor)}", 
                                font=('Arial', 14, 'bold'), bg='#ecf0f1')
        self.honor_label.pack(pady=5)
        
        self.jasa_sesungguhnya_label = tk.Label(self.summary_frame, 
                                            text=f"Jasa Sesungguhnya: {FormatUtils.format_currency(jasa_sesungguhnya)}", 
                                            font=('Arial', 14, 'bold'), bg='#ecf0f1', fg='#e74c3c')
        self.jasa_sesungguhnya_label.pack(pady=5)
        
        self.sekolah_label = tk.Label(self.summary_frame, 
                                    text=f"SEKOLAH: {self.processor.nama_sekolah}", 
                                    font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        self.sekolah_label.pack(anchor='w', pady=(10, 5))

    # Navigation methods
    def show_belanja_persediaan(self):
        if not self.processor.belanja_persediaan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_standard_results(self.processor.belanja_persediaan_items, 
                                    "Rincian Belanja Persediaan (5.1.02.01)", 
                                    "Total Belanja Persediaan")

    def show_belanja_jasa(self):
        if not self.processor.belanja_jasa_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_belanja_jasa_results(self.processor.belanja_jasa_items)

    def show_belanja_pemeliharaan(self):
        if not self.processor.belanja_pemeliharaan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_standard_results(self.processor.belanja_pemeliharaan_items,
                                    "Rincian Belanja Pemeliharaan (5.1.02.03)",
                                    "Total Belanja Pemeliharaan")

    def show_belanja_perjalanan(self):
        if not self.processor.belanja_perjalanan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_standard_results(self.processor.belanja_perjalanan_items,
                                    "Rincian Perjalanan Dinas (5.1.02.04)",
                                    "Total Belanja Perjalanan Dinas")

    def show_peralatan(self):
        if not self.processor.peralatan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_standard_results(self.processor.peralatan_items,
                                    "Rincian Peralatan dan Mesin (5.2.02)",
                                    "Total Peralatan")

    def show_aset_tetap(self):
        if not self.processor.aset_tetap_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_standard_results(self.processor.aset_tetap_items,
                                    "Rincian Aset Tetap Lainnya (5.2.04 & 5.2.05)",
                                    "Total Aset Tetap Lainnya")

    def show_ringkasan(self):
        if not self.processor.excel_data and self.processor.total_penerimaan == 0:
            messagebox.showwarning("Peringatan", "File belum diupload!")
            return
        
        # Create summary table
        columns = ('Kategori', 'Jumlah (Rp)')
        self.create_standard_table("Ringkasan Anggaran", columns)
        
        # Get summary data from processor
        summary_data = self.processor.get_summary_data()
        
        # Data ringkasan dengan background color info
        ringkasan_data = [
            ("PAGU TAHUN 2025", self.processor.total_penerimaan, True),
            ("BELANJA OPERASI", summary_data['total_belanja_persediaan'], True),
            ("  BELANJA HONOR", summary_data['total_honor'], False),
            ("  BELANJA JASA", summary_data['jasa_sesungguhnya'], False),
            ("  BELANJA PEMELIHARAAN", summary_data['total_pemeliharaan'], False),
            ("  BELANJA PERJALANAN", summary_data['total_perjalanan'], False),
            ("  BELANJA PERSEDIAAN", summary_data['belanja_persediaan_ringkasan'], False),
            ("BELANJA MODAL", summary_data['belanja_modal'], True),
            ("  PERALATAN DAN MESIN", summary_data['total_peralatan'], False),
            ("  ASET TETAP LAINNYA", summary_data['total_aset_tetap'], False),
            ("TOTAL ANGGARAN", summary_data['total_anggaran'], True)
        ]
        
        # Insert data into table
        for kategori, jumlah, is_highlight in ringkasan_data:
            formatted_jumlah = FormatUtils.format_currency(jumlah)
            item_id = self.tree.insert('', 'end', values=(kategori, formatted_jumlah))
            
            # Set background color for highlighted items
            if is_highlight:
                self.tree.set(item_id, 'Kategori', kategori)
                self.tree.set(item_id, 'Jumlah (Rp)', formatted_jumlah)
                # Configure tag for green background
                self.tree.tag_configure('highlight', background='#2ecc71', foreground='white')
                self.tree.item(item_id, tags=('highlight',))
        
        # Add school name label
        sekolah_label = tk.Label(self.summary_frame, text=f"SEKOLAH: {self.processor.nama_sekolah}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        sekolah_label.pack(anchor='w', pady=5)

    # Canvas scrolling event handlers
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling untuk scroll horizontal"""
        # Windows dan MacOS
        if event.delta:
            self.button_canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        # Linux
        elif event.num == 4:
            self.button_canvas.xview_scroll(-1, "units")
        elif event.num == 5:
            self.button_canvas.xview_scroll(1, "units")
        
    def _on_canvas_configure(self, event):
        """Handle canvas configure event"""
        # Update scroll region
        self.button_canvas.configure(scrollregion=self.button_canvas.bbox("all"))
        
        # Update canvas window height to match canvas
        canvas_height = self.button_canvas.winfo_height()
        self.button_canvas.itemconfig(self.canvas_window, height=canvas_height)
    
    def _on_frame_configure(self, event):
        """Handle frame configure event"""
        # Update scroll region when frame changes
        self.button_canvas.configure(scrollregion=self.button_canvas.bbox("all"))