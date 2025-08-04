"""
Main GUI application for SIKELAR
Handles user interface and interactions with enhanced scrolling
Modified to support split canvas layout for RKAS and BKU realisasi
Added dropdown for Triwulan selection in BKU section
Updated with aligned layout and auto BKU display for Belanja Persediaan and Pemeliharaan
Added RKAS placeholder functionality
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
from .base_page import BasePage  # Import BasePage

class RKASPage(BasePage):  # Inherit dari BasePage
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app)  # Call parent constructor
        
        # Initialize data processor
        self.processor = BOSDataProcessor()
        
        # For supporting active button highlight
        self.tab_buttons = {}
        self.active_tab = None
        
        # Triwulan selection variable
        self.selected_triwulan = tk.StringVar(value="Triwulan 1")
        
        # Canvas references
        self.canvas = None
        self.scrollbar = None
        self.scrollable_frame = None

    def build_page(self):
        """Build the RKAS page content - implemented from BasePage"""
        self.page_frame.configure(bg='#f8f9fa')
        
        # Create main canvas and scrollbar for scrolling
        self.canvas = tk.Canvas(self.page_frame, bg='#f8f9fa', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.page_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f8f9fa')
        
        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window in canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Configure canvas window to resize with canvas
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Bind mouse wheel for scrolling
        self.bind_mousewheel()
        
        # Build the actual page content
        self.setup_ui()

    def on_show(self):
        """Called when RKAS page is shown"""
        # Bind mousewheel events
        self.bind_mousewheel()
        
        # Force update canvas after showing
        if self.canvas:
            self.main_app.root.after(10, self.force_update_canvas)

    def on_hide(self):
        """Called when RKAS page is hidden"""
        # Unbind mousewheel events
        self.unbind_mousewheel()

    def force_update_canvas(self):
        """Force update canvas size and scroll region"""
        if self.canvas and self.canvas.winfo_exists():
            self.canvas.update_idletasks()
            self.scrollable_frame.update_idletasks()
            
            canvas_width = self.canvas.winfo_width()
            if canvas_width > 1:
                self.canvas.itemconfig(self.canvas_window, width=canvas_width)
            
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def bind_mousewheel(self):
        """Bind mouse wheel events for scrolling"""
        def on_mousewheel(event):
            if self.canvas and self.canvas.winfo_exists():
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        if self.canvas:
            self.canvas.bind("<MouseWheel>", on_mousewheel)
            self.bind_mousewheel_recursive(self.scrollable_frame, on_mousewheel)

    def bind_mousewheel_recursive(self, widget, callback):
        """Recursively bind mousewheel to all child widgets"""
        try:
            widget.bind("<MouseWheel>", callback, add='+')
            for child in widget.winfo_children():
                self.bind_mousewheel_recursive(child, callback)
        except tk.TclError:
            pass

    def unbind_mousewheel(self):
        """Unbind mouse wheel events"""
        try:
            if self.canvas and self.canvas.winfo_exists():
                self.canvas.unbind("<MouseWheel>")
            if self.scrollable_frame and self.scrollable_frame.winfo_exists():
                self.unbind_mousewheel_recursive(self.scrollable_frame)
        except tk.TclError:
            pass

    def unbind_mousewheel_recursive(self, widget):
        """Recursively unbind mousewheel from all child widgets"""
        try:
            widget.unbind("<MouseWheel>")
            for child in widget.winfo_children():
                self.unbind_mousewheel_recursive(child)
        except tk.TclError:
            pass

    def on_canvas_configure(self, event):
        """Handle canvas resize events"""
        if self.canvas and self.canvas.winfo_exists():
            canvas_width = event.width
            if canvas_width > 1:
                self.canvas.itemconfig(self.canvas_window, width=canvas_width)
                self.main_app.root.after_idle(self.force_update_canvas)

    def setup_ui(self):
        """Setup the main user interface - diubah untuk menggunakan scrollable_frame"""
        self._create_header()
        self._create_upload_section() 
        self._create_button_section()
        self._create_split_results_section()

    def _create_header(self):
        """Create header section with back button"""
        header_frame = tk.Frame(self.scrollable_frame, bg='#4a69bd', height=100)
        header_frame.pack(fill='x', pady=10)
        header_frame.pack_propagate(False)
        
        # Back button positioned at left
        back_btn = tk.Button(
            header_frame,
            text="🏠 Home",
            font=("Segoe UI", 11, "bold"),
            bg='#3742a6',
            fg='white',
            padx=25,
            pady=10,
            border=0,
            cursor='hand2',
            relief='flat',
            command=self.main_app.show_home_page
        )
        back_btn.place(x=25, y=22)
        
        # Add hover effects for back button
        back_btn.bind("<Enter>", lambda e: self.on_hover(e, '#2c3393'))
        back_btn.bind("<Leave>", lambda e: self.on_leave(e, '#3742a6'))
        
        # Centered title container
        title_container = tk.Frame(header_frame, bg='#4a69bd')
        title_container.place(relx=0.5, rely=0.5, anchor='center')
        
        # Main title
        header_label = tk.Label(title_container, 
                               text="Sistem Perhitungan Anggaran RKAS dan Realisasi BKU",
                               font=("Segoe UI", 18, "bold"), 
                               bg='#4a69bd', 
                               fg='white')
        header_label.pack()
        
        # Subtitle
        subtitle_label = tk.Label(title_container,
                                 text="Upload File Excel dan Analisis Data Anggaran",
                                 font=("Segoe UI", 11),
                                 bg='#4a69bd',
                                 fg='#c8d4f0')
        subtitle_label.pack()

    def on_hover(self, event, hover_color):
        """Event handler for button hover"""
        event.widget.config(bg=hover_color)

    def on_leave(self, event, original_color):
        """Event handler for button leave"""
        event.widget.config(bg=original_color)

    def _create_upload_section(self):
        """Create upload section - diubah untuk menggunakan scrollable_frame"""
        upload_frame = tk.Frame(self.scrollable_frame, bg='#ecf0f1', relief='raised', bd=2)
        upload_frame.pack(fill='x', padx=10, pady=10)
        
        upload_label = tk.Label(upload_frame, text="Upload RKAS dan BKU Dalam 1 File Excel :", 
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
        """Create navigation button section - diubah untuk menggunakan scrollable_frame"""
        main_button_frame = tk.Frame(self.scrollable_frame, bg='#f0f0f0', height=80)
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

    def _create_split_results_section(self):
        """Create split results display section - diubah untuk menggunakan scrollable_frame"""
        # Main container for split results
        self.main_results_container = tk.Frame(self.scrollable_frame, bg='#f0f0f0')
        self.main_results_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Left side - RKAS
        self._create_rkas_section()
        
        # Separator
        separator = tk.Frame(self.main_results_container, bg='#bdc3c7', width=2)
        separator.pack(side='left', fill='y', padx=5)
        
        # Right side - BKU Realisasi
        self._create_bku_section()

    def _create_rkas_section(self):
        """Create RKAS section on the left"""
        # RKAS Container
        self.rkas_container = tk.Frame(self.main_results_container, bg='#ffffff', relief='sunken', bd=2)
        self.rkas_container.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        # RKAS Header
        rkas_header = tk.Label(self.rkas_container, text="  RKAS (Rencana Kegiatan dan Anggaran Sekolah)", 
                            font=('Arial', 14, 'bold'), bg='#3498db', fg='white', pady=10, 
                            anchor='w')
        rkas_header.pack(fill='x')
        
        # RKAS Canvas for scrolling
        self.rkas_canvas = tk.Canvas(self.rkas_container, bg='#ffffff', highlightthickness=0)
        self.rkas_canvas.pack(side='left', fill='both', expand=True)
        
        # RKAS Vertical scrollbar
        self.rkas_v_scrollbar = tk.Scrollbar(self.rkas_container, orient='vertical', 
                                           command=self.rkas_canvas.yview)
        self.rkas_v_scrollbar.pack(side='right', fill='y')
        self.rkas_canvas.configure(yscrollcommand=self.rkas_v_scrollbar.set)
        
        # RKAS Scrollable frame
        self.rkas_frame = tk.Frame(self.rkas_canvas, bg='#ffffff')
        self.rkas_canvas_window = self.rkas_canvas.create_window((0, 0), window=self.rkas_frame, anchor='nw')
        
        # Bind RKAS events
        self.rkas_frame.bind('<Configure>', self._on_rkas_frame_configure)
        self.rkas_canvas.bind('<Configure>', self._on_rkas_canvas_configure)
        self.rkas_canvas.bind("<MouseWheel>", self._on_rkas_mousewheel)
        self.rkas_canvas.bind("<Button-4>", self._on_rkas_mousewheel)
        self.rkas_canvas.bind("<Button-5>", self._on_rkas_mousewheel)
        
        # Create initial RKAS placeholder
        self._create_rkas_placeholder()

    def _create_rkas_placeholder(self):
        """Create placeholder content for RKAS section"""
        # Check if RKAS data is available
        if hasattr(self.processor, 'excel_data') and self.processor.excel_data:
            # If data is available but no specific category is selected
            instruction_label = tk.Label(self.rkas_frame, 
                                    text="Pilih kategori anggaran\nuntuk melihat data RKAS", 
                                    font=('Arial', 14, 'bold'), 
                                    fg='#2c3e50', bg='#ffffff', pady=30)
            instruction_label.pack(expand=True)
            
            detail_label = tk.Label(self.rkas_frame, 
                                text="Data RKAS tersedia.\nKlik salah satu tombol kategori di atas.", 
                                font=('Arial', 12), 
                                fg='#7f8c8d', bg='#ffffff')
            detail_label.pack(pady=10)
        else:
            # Original placeholder when no RKAS data
            placeholder_label = tk.Label(self.rkas_frame, 
                                    text="Fitur RKAS\n(Upload File Terlebih Dahulu, Lalu Pilih Kategori)", 
                                    font=('Arial', 16, 'italic'), 
                                    fg='#7f8c8d', bg='#ffffff', pady=50)
            placeholder_label.pack(expand=True)

    def _create_bku_section(self):
        """Create BKU Realisasi section on the right with Triwulan dropdown"""
        # BKU Container
        self.bku_container = tk.Frame(self.main_results_container, bg='#ffffff', relief='sunken', bd=2)
        self.bku_container.pack(side='right', fill='both', expand=True, padx=(5, 0))
        
        # BKU Header Frame (contains title and dropdown)
        bku_header_frame = tk.Frame(self.bku_container, bg='#e74c3c')
        bku_header_frame.pack(fill='x')
        
        # BKU Title (left side of header)
        bku_title = tk.Label(bku_header_frame, text="BKU (Buku Kas Umum) - Realisasi", 
                            font=('Arial', 14, 'bold'), bg='#e74c3c', fg='white')
        bku_title.pack(side='left', padx=10, pady=10)
        
        # Triwulan Dropdown Frame (right side of header)
        dropdown_frame = tk.Frame(bku_header_frame, bg='#e74c3c')
        dropdown_frame.pack(side='right', padx=10, pady=10)
        
        # Triwulan Dropdown
        self.triwulan_combobox = ttk.Combobox(dropdown_frame, 
                                             textvariable=self.selected_triwulan,
                                             values=["Triwulan 1", "Triwulan 2", "Triwulan 3", "Triwulan 4"],
                                             state="readonly",
                                             font=('Arial', 10, 'bold'),
                                             width=12)
        self.triwulan_combobox.pack(side='right')
        
        # Bind dropdown selection event
        self.triwulan_combobox.bind('<<ComboboxSelected>>', self.on_triwulan_changed)
        
        # BKU Canvas for scrolling
        self.bku_canvas = tk.Canvas(self.bku_container, bg='#ffffff', highlightthickness=0)
        self.bku_canvas.pack(side='left', fill='both', expand=True)
        
        # BKU Vertical scrollbar
        self.bku_v_scrollbar = tk.Scrollbar(self.bku_container, orient='vertical', 
                                          command=self.bku_canvas.yview)
        self.bku_v_scrollbar.pack(side='right', fill='y')
        self.bku_canvas.configure(yscrollcommand=self.bku_v_scrollbar.set)
        
        # BKU Scrollable frame
        self.bku_frame = tk.Frame(self.bku_canvas, bg='#ffffff')
        self.bku_canvas_window = self.bku_canvas.create_window((0, 0), window=self.bku_frame, anchor='nw')
        
        # Bind BKU events
        self.bku_frame.bind('<Configure>', self._on_bku_frame_configure)
        self.bku_canvas.bind('<Configure>', self._on_bku_canvas_configure)
        self.bku_canvas.bind("<MouseWheel>", self._on_bku_mousewheel)
        self.bku_canvas.bind("<Button-4>", self._on_bku_mousewheel)
        self.bku_canvas.bind("<Button-5>", self._on_bku_mousewheel)
        
        # Placeholder content for BKU section
        self._create_bku_placeholder()

    def _create_bku_placeholder(self):
        """Create placeholder content for BKU section - Updated"""
        # Check if BKU data is available
        if hasattr(self.processor, 'bku_data_available') and self.processor.bku_data_available:
            # Check if current active tab supports BKU display
            current_tab = self.active_tab or "Belanja Persediaan"
            if current_tab in ["Belanja Persediaan", "Pemeliharaan", "Perjalanan Dinas", "Peralatan", "Aset Tetap", "Jasa"]:
                # Show instruction to select triwulan
                instruction_label = tk.Label(self.bku_frame, 
                                        text="Pilih Triwulan untuk melihat\ndata realisasi BKU", 
                                        font=('Arial', 14, 'bold'), 
                                        fg='#2c3e50', bg='#ffffff', pady=30)
                instruction_label.pack(expand=True)
                
                detail_label = tk.Label(self.bku_frame, 
                                    text="Data BKU tersedia.\nGunakan dropdown di atas untuk memilih periode.", 
                                    font=('Arial', 12), 
                                    fg='#7f8c8d', bg='#ffffff')
                detail_label.pack(pady=10)
            else:
                # Show message for non-supported categories
                placeholder_label = tk.Label(self.bku_frame, 
                                        text="Fitur Realisasi BKU\nTidak tersedia untuk kategori ini", 
                                        font=('Arial', 14, 'italic'), 
                                        fg='#7f8c8d', bg='#ffffff', pady=50)
                placeholder_label.pack(expand=True)
        else:
            # Original placeholder when no BKU data
            placeholder_label = tk.Label(self.bku_frame, 
                                    text="Fitur Realisasi BKU\n(Upload File Terlebih Dahulu, Lalu Pilih Kategori)", 
                                    font=('Arial', 16, 'italic'), 
                                    fg='#7f8c8d', bg='#ffffff', pady=50)
            placeholder_label.pack(expand=True)

    def on_triwulan_changed(self, event=None):
        """Handle triwulan dropdown selection change - UPDATED untuk include ringkasan"""
        selected = self.selected_triwulan.get()
        print(f"Triwulan dipilih: {selected}")
        
        # Check if we're currently on Ringkasan tab
        if self.active_tab == "Ringkasan":
            # Update BKU summary display
            if hasattr(self.processor, 'bku_data_available') and self.processor.bku_data_available:
                self._display_bku_summary_for_triwulan(selected)
            return
        
        # Existing code for other categories
        if self.active_tab in ["Belanja Persediaan", "Pemeliharaan", "Perjalanan Dinas", "Peralatan", "Aset Tetap", "Jasa"]:
            category_map = {
                "Belanja Persediaan": "Belanja Persediaan",
                "Pemeliharaan": "Pemeliharaan",
                "Perjalanan Dinas": "Perjalanan Dinas", 
                "Peralatan": "Peralatan",
                "Aset Tetap": "Aset Tetap",
                "Jasa": "Jasa"
            }
            self._display_bku_for_category(category_map[self.active_tab])

    def _display_bku_for_category(self, category):
        """Display BKU data for specific category with appropriate summary format"""
        # Clear current BKU content
        for widget in self.bku_frame.winfo_children():
            widget.destroy()
            
        # Reset BKU canvas scroll position to top
        self.bku_canvas.yview_moveto(0)
        
        # Check if BKU data is available
        if not hasattr(self.processor, 'bku_data_available') or not self.processor.bku_data_available:
            self._create_bku_placeholder()
            return
        
        # Set default triwulan and get data
        selected_triwulan = self.selected_triwulan.get()
        
        # Get BKU data based on category dan panggil method yang sesuai
        if category == "Jasa":
            bku_data = self.processor.get_bku_belanja_jasa_by_triwulan(selected_triwulan)
            title = f"Realisasi Belanja Jasa (5.1.02.02) - {selected_triwulan}"
            if bku_data:
                # KHUSUS JASA - gunakan method dengan breakdown honor
                self._display_bku_jasa_data(selected_triwulan, bku_data, title)
            else:
                # No data handling
                no_data_label = tk.Label(self.bku_frame, 
                                    text=f"Tidak ada data realisasi\nuntuk {selected_triwulan}", 
                                    font=('Arial', 14, 'bold'), 
                                    fg='#e74c3c', bg='#ffffff', pady=30)
                no_data_label.pack(expand=True)
                self._show_triwulan_status(selected_triwulan)
        elif category == "Belanja Persediaan":
            bku_data = self.processor.get_bku_belanja_persediaan_by_triwulan(selected_triwulan)
            title = f"Realisasi Belanja Persediaan (5.1.02.01) - {selected_triwulan}"
            if bku_data:
                # KATEGORI LAIN - gunakan method generic tanpa breakdown honor
                self._display_bku_data_generic(selected_triwulan, bku_data, title)
            else:
                self._handle_no_bku_data(selected_triwulan)
        elif category == "Pemeliharaan":
            bku_data = self.processor.get_bku_belanja_pemeliharaan_by_triwulan(selected_triwulan)
            title = f"Realisasi Belanja Pemeliharaan (5.1.02.03) - {selected_triwulan}"
            if bku_data:
                self._display_bku_data_generic(selected_triwulan, bku_data, title)
            else:
                self._handle_no_bku_data(selected_triwulan)
        elif category == "Perjalanan Dinas":
            bku_data = self.processor.get_bku_belanja_perjalanan_by_triwulan(selected_triwulan)
            title = f"Realisasi Belanja Perjalanan Dinas (5.1.02.04) - {selected_triwulan}"
            if bku_data:
                self._display_bku_data_generic(selected_triwulan, bku_data, title)
            else:
                self._handle_no_bku_data(selected_triwulan)
        elif category == "Peralatan":
            bku_data = self.processor.get_bku_peralatan_by_triwulan(selected_triwulan)
            title = f"Realisasi Peralatan dan Mesin (5.2.02) - {selected_triwulan}"
            if bku_data:
                self._display_bku_data_generic(selected_triwulan, bku_data, title)
            else:
                self._handle_no_bku_data(selected_triwulan)
        elif category == "Aset Tetap":
            bku_data = self.processor.get_bku_aset_tetap_by_triwulan(selected_triwulan)
            title = f"Realisasi Aset Tetap Lainnya (5.2.04 & 5.2.05) - {selected_triwulan}"
            if bku_data:
                self._display_bku_data_generic(selected_triwulan, bku_data, title)
            else:
                self._handle_no_bku_data(selected_triwulan)
        else:
            # For other categories, show placeholder
            self._create_bku_placeholder()
            return
        
        # Update canvas scroll region
        self.bku_frame.update_idletasks()
        self.bku_canvas.configure(scrollregion=self.bku_canvas.bbox("all"))

    def _handle_no_bku_data(self, selected_triwulan):
        """Handle case when no BKU data is available for selected triwulan"""
        no_data_label = tk.Label(self.bku_frame, 
                            text=f"Tidak ada data realisasi\nuntuk {selected_triwulan}", 
                            font=('Arial', 14, 'bold'), 
                            fg='#e74c3c', bg='#ffffff', pady=30)
        no_data_label.pack(expand=True)
        
        # Show triwulan status
        self._show_triwulan_status(selected_triwulan)

    # RKAS Canvas Event Handlers
    def _on_rkas_frame_configure(self, event):
        """Handle RKAS frame configure event for vertical scrolling"""
        self.rkas_canvas.configure(scrollregion=self.rkas_canvas.bbox("all"))

    def _on_rkas_canvas_configure(self, event):
        """Handle RKAS canvas configure event"""
        canvas_width = self.rkas_canvas.winfo_width()
        self.rkas_canvas.itemconfig(self.rkas_canvas_window, width=canvas_width)

    def _on_rkas_mousewheel(self, event):
        """Handle mouse wheel scrolling for RKAS area"""
        if event.delta:
            self.rkas_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        elif event.num == 4:
            self.rkas_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.rkas_canvas.yview_scroll(1, "units")

    # BKU Canvas Event Handlers
    def _on_bku_frame_configure(self, event):
        """Handle BKU frame configure event for vertical scrolling"""
        self.bku_canvas.configure(scrollregion=self.bku_canvas.bbox("all"))

    def _on_bku_canvas_configure(self, event):
        """Handle BKU canvas configure event"""
        canvas_width = self.bku_canvas.winfo_width()
        self.bku_canvas.itemconfig(self.bku_canvas_window, width=canvas_width)

    def _on_bku_mousewheel(self, event):
        """Handle mouse wheel scrolling for BKU area"""
        if event.delta:
            self.bku_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        elif event.num == 4:
            self.bku_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.bku_canvas.yview_scroll(1, "units")

    def upload_excel(self):
        """Handle Excel file upload - UPDATED to process BKU"""
        file_path = filedialog.askopenfilename(title="Pilih File Excel RKAS", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.processor.extract_excel_data(file_path)
                self.file_label.config(text=f"File dipilih: {file_path.split('/')[-1]}", fg='#27ae60')

                # Disable upload button after successful upload
                self.upload_btn.config(state='disabled')

                # Update RKAS placeholder to show available data
                self._update_rkas_placeholder_after_upload()

                # Update BKU placeholder to show available data
                self._create_bku_placeholder()

                # Info message dengan detail sheet
                bku_status = "BKU sheet ditemukan dan diproses" if self.processor.bku_data_available else "BKU sheet tidak ditemukan"
                
                # Count BKU data if available
                bku_info = ""
                if self.processor.bku_data_available:
                    total_bku_items = 0
                    for triwulan in ['Triwulan 1', 'Triwulan 2', 'Triwulan 3', 'Triwulan 4']:
                        items = self.processor.get_bku_belanja_persediaan_by_triwulan(triwulan)
                        if items:
                            total_bku_items += len(items)
                            bku_info += f"\n{triwulan}: {len(items)} item realisasi"
                    
                    if total_bku_items > 0:
                        bku_status += f" ({total_bku_items} total item){bku_info}"
                
                messagebox.showinfo("Berhasil", 
                    f"File Excel berhasil diproses!\n"
                    f"Sheet RKAS: Berhasil diproses\n"
                    f"Sheet BKU: Berhasil diproses\n"
                    f"Total Penerimaan: {FormatUtils.format_currency(self.processor.total_penerimaan)}\n"
                    f"Ditemukan {len(self.processor.belanja_persediaan_items)} item belanja persediaan\n"
                    f"Ditemukan {len(self.processor.belanja_jasa_items)} item belanja jasa\n"
                    f"Ditemukan {len(self.processor.belanja_pemeliharaan_items)} item belanja pemeliharaan\n"
                    f"Ditemukan {len(self.processor.belanja_perjalanan_items)} item belanja perjalanan\n"
                    f"Ditemukan {len(self.processor.peralatan_items)} item peralatan dan mesin\n"
                    f"Ditemukan {len(self.processor.aset_tetap_items)} item aset tetap lainnya")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal membaca file Excel: {str(e)}")

    def _update_rkas_placeholder_after_upload(self):
        """Update RKAS placeholder after successful file upload"""
        # Clear current RKAS content
        for widget in self.rkas_frame.winfo_children():
            widget.destroy()
        
        # Reset RKAS canvas scroll position to top
        self.rkas_canvas.yview_moveto(0)
        
        # Create updated placeholder
        self._create_rkas_placeholder()
        
        # Update canvas scroll region
        self.rkas_frame.update_idletasks()
        self.rkas_canvas.configure(scrollregion=self.rkas_canvas.bbox("all"))
    
    def reset_data(self):
        """Reset all data and UI - UPDATED"""
        self.processor.reset_data()

        # Clear RKAS display and recreate placeholder
        for widget in self.rkas_frame.winfo_children():
            widget.destroy()
        
        # Re-add RKAS placeholder
        self._create_rkas_placeholder()

        # Clear BKU display and recreate placeholder
        for widget in self.bku_frame.winfo_children():
            widget.destroy()
        
        # Re-add BKU placeholder
        self._create_bku_placeholder()
        
        # Reset triwulan selection to default
        self.selected_triwulan.set("Triwulan 1")

        # Reset scroll positions
        self.rkas_canvas.yview_moveto(0)
        self.bku_canvas.yview_moveto(0)

        # Clear file label
        self.file_label.config(text="Belum ada file yang dipilih")

        # Re-enable upload button
        self.upload_btn.config(state='normal')

        # Reset active tab
        self.active_tab = None
        
        # Reset button highlights
        for tab_name, (btn, orig_color) in self.tab_buttons.items():
            btn.config(bg=orig_color)

        messagebox.showinfo("Reset", "Data berhasil dibersihkan.")

    def create_standard_table(self, title, columns):
        """Create standard table layout in RKAS section with aligned positioning"""
        for widget in self.rkas_frame.winfo_children():
            widget.destroy()
        
        # Reset canvas scroll position to top
        self.rkas_canvas.yview_moveto(0)
        
        # FIXED: Add consistent top padding for title alignment
        title_frame = tk.Frame(self.rkas_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(20, 10))  # Same padding as BKU section
        
        self.result_title = tk.Label(title_frame, text=title, 
                                    font=('Arial', 14, 'bold'), bg='#ffffff', fg='#2c3e50')
        self.result_title.pack()
        
        self.table_frame = tk.Frame(self.rkas_frame, bg='#ffffff')
        self.table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Adjust table height for split view
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show='headings', height=15)
        
        # Style untuk memperbesar font tabel
        style = ttk.Style()
        style.configure("Treeview", font=('Arial', 10))
        style.configure("Treeview.Heading", font=('Arial', 12, 'bold'))
        
        for col in columns:
            self.tree.heading(col, text=col)
        
        # Adjust column widths for split view
        if len(columns) == 4:  # Standard detail table
            self.tree.column(columns[0], width=120, anchor='center')  # Kode Rekening
            self.tree.column(columns[1], width=100, anchor='center')  # Kode Kegiatan
            self.tree.column(columns[2], width=280, anchor='w')       # Uraian
            self.tree.column(columns[3], width=120, anchor='w')       # Jumlah
        elif len(columns) == 2:  # Summary table
            self.tree.column(columns[0], width=320, anchor='w')       # Kategori
            self.tree.column(columns[1], width=160, anchor='w')       # Jumlah
        
        table_scrollbar = ttk.Scrollbar(self.table_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=table_scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        table_scrollbar.pack(side='right', fill='y')
        
        self.summary_frame = tk.Frame(self.rkas_frame, bg='#ecf0f1', relief='raised', bd=1)
        self.summary_frame.pack(fill='x', padx=10, pady=10)
        
        # Force update of canvas scroll region
        self.rkas_frame.update_idletasks()
        self.rkas_canvas.configure(scrollregion=self.rkas_canvas.bbox("all"))

    def _display_bku_data_generic(self, triwulan, data, title):
        """Display BKU realisasi data in table format - TANPA BREAKDOWN HONOR untuk kategori selain Jasa"""
        # Title frame
        title_frame = tk.Frame(self.bku_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(20, 10))
        
        title_label = tk.Label(title_frame, 
                            text=title, 
                            font=('Arial', 14, 'bold'), 
                            bg='#ffffff', fg='#2c3e50')
        title_label.pack()
        
        # Create table frame
        table_frame = tk.Frame(self.bku_frame, bg='#ffffff')
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create treeview for BKU data
        columns = ('Tanggal', 'Kode Rekening', 'Kode Kegiatan', 'Uraian', 'Jumlah (Rp)')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configure column headings and widths
        for col in columns:
            tree.heading(col, text=col)
        
        tree.column('Tanggal', width=80, anchor='center')
        tree.column('Kode Rekening', width=100, anchor='center')
        tree.column('Kode Kegiatan', width=80, anchor='center')
        tree.column('Uraian', width=250, anchor='w')
        tree.column('Jumlah (Rp)', width=120, anchor='w')
        
        # Insert data - TANPA LOGIK HONOR untuk kategori selain Jasa
        total_realisasi = 0
        for item in data:
            formatted_tanggal = item['tanggal'].strftime('%d-%m-%Y')
            formatted_jumlah = FormatUtils.format_currency(item['jumlah'])
            
            tree.insert('', 'end', values=(
                formatted_tanggal,
                item['kode_rekening'],
                item['kode_kegiatan'],
                item['uraian'],
                formatted_jumlah
            ))
            total_realisasi += item['jumlah']
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Summary section - SIMPLE SUMMARY untuk kategori selain Jasa
        summary_frame = tk.Frame(self.bku_frame, bg='#ecf0f1', relief='raised', bd=1)
        summary_frame.pack(fill='x', padx=10, pady=10)
        
        # Total realisasi saja - TANPA BREAKDOWN
        total_label = tk.Label(summary_frame, 
                            text=f"Total Realisasi {triwulan}: {FormatUtils.format_currency(total_realisasi)}", 
                            font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#27ae60')
        total_label.pack(pady=5)
        
        # School name
        school_label = tk.Label(summary_frame, 
                            text=f"SEKOLAH {self.processor.nama_sekolah}", 
                            font=('Arial', 10, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        school_label.pack(anchor='w', pady=(10, 5))


    def _show_triwulan_status(self, triwulan):
        """Show status information for triwulan"""
        # Determine months for triwulan
        if triwulan == "Triwulan 1":
            months = "Januari - Maret"
        elif triwulan == "Triwulan 2":
            months = "April - Juni"
        elif triwulan == "Triwulan 3":
            months = "Juli - September"
        elif triwulan == "Triwulan 4":
            months = "Oktober - Desember"
        else:
            months = "Tidak dikenal"
        
        info_label = tk.Label(self.bku_frame, 
                            text=f"Periode: {months}\n\nKemungkinan penyebab:\n• Data {triwulan} belum lengkap\n• Bulan terakhir periode belum ada data\n• Kode rekening tidak ditemukan", 
                            font=('Arial', 11), 
                            fg='#7f8c8d', bg='#ffffff',
                            justify='left')
        info_label.pack(pady=20, padx=20)

    def display_standard_results(self, items: List[Dict], title: str, total_label: str):
        """Display results in standard format with consistent summary layout"""
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
        
        # Create consistent summary layout
        self._create_consistent_summary(total_label, total_jumlah)

    def _create_consistent_summary(self, total_label: str, total_jumlah: int):
        """Create consistent summary section with center total and left school name - TANPA BREAKDOWN HONOR"""
        # Main total label at center
        self.total_label = tk.Label(self.summary_frame, text=f"{total_label}: {FormatUtils.format_currency(total_jumlah)}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#27ae60')
        self.total_label.pack(pady=5)
        
        # School name label at bottom left
        self.sekolah_label = tk.Label(self.summary_frame, 
                                    text=f"SEKOLAH {self.processor.nama_sekolah}", 
                                    font=('Arial', 10, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        self.sekolah_label.pack(anchor='w', pady=(10, 5))

    def display_belanja_jasa_results(self, items: List[Dict]):
        """Display belanja jasa results with extended summary including honor breakdown"""
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
        
        # Calculate honor and actual service - KHUSUS UNTUK JASA
        honor_items = self.processor.filter_budget_by_codes(self.processor.kategori_kode['honor'])
        total_honor = sum(item['jumlah'] for item in honor_items)
        jasa_sesungguhnya = total_belanja_jasa - total_honor
        
        # Create extended summary dengan data honor breakdown - HANYA UNTUK JASA
        self.total_label = tk.Label(self.summary_frame, text=f"Total Belanja Jasa (RKAS): {FormatUtils.format_currency(total_belanja_jasa)}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1')
        self.total_label.pack(pady=5)
        
        self.honor_label = tk.Label(self.summary_frame, 
                                text=f"Pembayaran Honor (RKAS): {FormatUtils.format_currency(total_honor)}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1')
        self.honor_label.pack(pady=5)
        
        self.jasa_sesungguhnya_label = tk.Label(self.summary_frame, 
                                            text=f"Jasa Sesungguhnya (RKAS): {FormatUtils.format_currency(jasa_sesungguhnya)}", 
                                            font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#27ae60')
        self.jasa_sesungguhnya_label.pack(pady=5)
        
        self.sekolah_label = tk.Label(self.summary_frame, 
                                    text=f"SEKOLAH {self.processor.nama_sekolah}", 
                                    font=('Arial', 10, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        self.sekolah_label.pack(anchor='w', pady=(10, 5))

    def _display_bku_jasa_data(self, triwulan, data, title):
        """Display BKU jasa data dengan summary yang menghitung honor dan jasa sesungguhnya - KHUSUS JASA"""
        # Title frame
        title_frame = tk.Frame(self.bku_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(20, 10))
        
        title_label = tk.Label(title_frame, 
                            text=title, 
                            font=('Arial', 14, 'bold'), 
                            bg='#ffffff', fg='#2c3e50')
        title_label.pack()
        
        # Create table frame
        table_frame = tk.Frame(self.bku_frame, bg='#ffffff')
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create treeview for BKU data
        columns = ('Tanggal', 'Kode Rekening', 'Kode Kegiatan', 'Uraian', 'Jumlah (Rp)')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configure column headings and widths
        for col in columns:
            tree.heading(col, text=col)
        
        tree.column('Tanggal', width=80, anchor='center')
        tree.column('Kode Rekening', width=100, anchor='center')
        tree.column('Kode Kegiatan', width=80, anchor='center')
        tree.column('Uraian', width=250, anchor='w')
        tree.column('Jumlah (Rp)', width=120, anchor='w')
        
        # Insert data dan hitung honor vs jasa sesungguhnya - LOGIK KHUSUS JASA
        total_realisasi = 0
        total_honor_bku = 0
        
        for item in data:
            formatted_tanggal = item['tanggal'].strftime('%d-%m-%Y')
            formatted_jumlah = FormatUtils.format_currency(item['jumlah'])
            
            tree.insert('', 'end', values=(
                formatted_tanggal,
                item['kode_rekening'],
                item['kode_kegiatan'],
                item['uraian'],
                formatted_jumlah
            ))
            
            total_realisasi += item['jumlah']
            
            # Hitung honor berdasarkan kode kegiatan yang dimulai dengan 07.12 - KHUSUS JASA
            if item['kode_kegiatan'].startswith('07.12'):
                total_honor_bku += item['jumlah']
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Summary section dengan breakdown honor dan jasa sesungguhnya - KHUSUS JASA
        summary_frame = tk.Frame(self.bku_frame, bg='#ecf0f1', relief='raised', bd=1)
        summary_frame.pack(fill='x', padx=10, pady=10)
        
        # Total realisasi
        total_label = tk.Label(summary_frame, 
                            text=f"Total Realisasi {triwulan}: {FormatUtils.format_currency(total_realisasi)}", 
                            font=('Arial', 12, 'bold'), bg='#ecf0f1')
        total_label.pack(pady=5)
        
        # Honor realisasi - KHUSUS JASA
        honor_bku_label = tk.Label(summary_frame, 
                                text=f"Pembayaran Honor (Realisasi): {FormatUtils.format_currency(total_honor_bku)}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1')
        honor_bku_label.pack(pady=5)
        
        # Jasa sesungguhnya realisasi - KHUSUS JASA
        jasa_sesungguhnya_bku = total_realisasi - total_honor_bku
        jasa_bku_label = tk.Label(summary_frame, 
                                text=f"Jasa Sesungguhnya (Realisasi): {FormatUtils.format_currency(jasa_sesungguhnya_bku)}", 
                                font=('Arial', 12, 'bold'), bg='#ecf0f1', fg='#27ae60')
        jasa_bku_label.pack(pady=5)
        
        # School name
        school_label = tk.Label(summary_frame, 
                            text=f"SEKOLAH {self.processor.nama_sekolah}", 
                            font=('Arial', 10, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        school_label.pack(anchor='w', pady=(10, 5))

    def _display_bku_summary_for_triwulan(self, triwulan):
        """Display BKU summary data untuk triwulan tertentu - FIXED VERSION"""
        # Clear current BKU content
        for widget in self.bku_frame.winfo_children():
            widget.destroy()
            
        # Reset BKU canvas scroll position to top
        self.bku_canvas.yview_moveto(0)
        
        # Check if BKU data is available
        if not hasattr(self.processor, 'bku_data_available') or not self.processor.bku_data_available:
            print("Debug: BKU data not available in _display_bku_summary_for_triwulan")
            self._create_bku_placeholder()
            return
        
        # Get BKU summary data
        try:
            summary_data = self.processor.get_bku_summary_data_by_triwulan(triwulan)
            print(f"Debug: Got BKU summary data: {summary_data}")
        except Exception as e:
            print(f"Debug: Error getting BKU summary data: {e}")
            self._create_bku_placeholder()
            return
        
        if not summary_data or summary_data.get('total_realisasi', 0) == 0:
            # No data for this triwulan
            print(f"Debug: No BKU data for {triwulan}")
            no_data_label = tk.Label(self.bku_frame, 
                                text=f"Tidak ada data realisasi\nuntuk {triwulan}", 
                                font=('Arial', 14, 'bold'), 
                                fg='#e74c3c', bg='#ffffff', pady=30)
            no_data_label.pack(expand=True)
            return
        
        print(f"Debug: Displaying BKU data with total realisasi: {summary_data['total_realisasi']}")
        
        # Title frame
        title_frame = tk.Frame(self.bku_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(20, 10))
        
        title_label = tk.Label(title_frame, 
                            text=f"Ringkasan Realisasi BKU - {triwulan}", 
                            font=('Arial', 14, 'bold'), 
                            bg='#ffffff', fg='#2c3e50')
        title_label.pack()
        
        # Create table frame
        table_frame = tk.Frame(self.bku_frame, bg='#ffffff')
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create treeview for summary
        columns = ('Kategori', 'Jumlah (Rp)')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configure column headings and widths
        for col in columns:
            tree.heading(col, text=col)
        
        tree.column('Kategori', width=320, anchor='w')
        tree.column('Jumlah (Rp)', width=160, anchor='w')
        
        # Data ringkasan BKU dengan background color info (mirip dengan RKAS)
        ringkasan_data = [
            ("BELANJA OPERASI", summary_data['total_belanja_operasi_bku'], True),
            ("  BELANJA HONOR", summary_data['total_honor_bku'], False),
            ("  BELANJA JASA", summary_data['jasa_sesungguhnya_bku'], False),
            ("  BELANJA PEMELIHARAAN", summary_data['total_pemeliharaan_bku'], False),
            ("  BELANJA PERJALANAN", summary_data['total_perjalanan_bku'], False),
            ("  BELANJA PERSEDIAAN", summary_data['total_persediaan_bku'], False),
            ("BELANJA MODAL", summary_data['belanja_modal_bku'], True),
            ("  PERALATAN DAN MESIN", summary_data['total_peralatan_bku'], False),
            ("  ASET TETAP LAINNYA", summary_data['total_aset_tetap_bku'], False),
            ("TOTAL REALISASI", summary_data['total_realisasi'], True)
        ]
        
        # Insert data into table
        for kategori, jumlah, is_highlight in ringkasan_data:
            formatted_jumlah = FormatUtils.format_currency(jumlah)
            item_id = tree.insert('', 'end', values=(kategori, formatted_jumlah))
            
            # Set background color for highlighted items
            if is_highlight:
                tree.set(item_id, 'Kategori', kategori)
                tree.set(item_id, 'Jumlah (Rp)', formatted_jumlah)
                # Configure tag for red background for BKU
                tree.tag_configure('highlight', background='#e74c3c', foreground='white')
                tree.item(item_id, tags=('highlight',))
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Summary section
        summary_frame = tk.Frame(self.bku_frame, bg='#ecf0f1', relief='raised', bd=1)
        summary_frame.pack(fill='x', padx=10, pady=10)
        
        # School name
        school_label = tk.Label(summary_frame, 
                            text=f"SEKOLAH {self.processor.nama_sekolah}", 
                            font=('Arial', 10, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        school_label.pack(anchor='w', pady=(10, 5))
        
        # Update canvas scroll region
        self.bku_frame.update_idletasks()
        self.bku_canvas.configure(scrollregion=self.bku_canvas.bbox("all"))


    # Navigation methods - UPDATED with auto BKU display for supported categories
    def show_belanja_persediaan(self):
        """Show both RKAS and BKU data for Belanja Persediaan"""
        if not self.processor.belanja_persediaan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        
        # Display RKAS data
        self.display_standard_results(self.processor.belanja_persediaan_items, 
                                    "Rincian Belanja Persediaan (5.1.02.01)", 
                                    "Total Belanja Persediaan")
        
        # AUTO-DISPLAY BKU DATA
        self._display_bku_for_category("Belanja Persediaan")

    def show_belanja_jasa(self):
        """Show RKAS data for Belanja Jasa"""
        if not self.processor.belanja_jasa_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_belanja_jasa_results(self.processor.belanja_jasa_items)
        
        # Clear BKU section for non-supported categories
        self._display_bku_for_category("Jasa")

    def show_belanja_pemeliharaan(self):
        """Show both RKAS and BKU data for Belanja Pemeliharaan"""
        if not self.processor.belanja_pemeliharaan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        
        # Display RKAS data
        self.display_standard_results(self.processor.belanja_pemeliharaan_items,
                                    "Rincian Belanja Pemeliharaan (5.1.02.03)",
                                    "Total Belanja Pemeliharaan")
        
        # AUTO-DISPLAY BKU DATA
        self._display_bku_for_category("Pemeliharaan")

    def show_belanja_perjalanan(self):
        """Show both RKAS and BKU data for Belanja Perjalanan"""
        if not self.processor.belanja_perjalanan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        
        # Display RKAS data
        self.display_standard_results(self.processor.belanja_perjalanan_items,
                                    "Rincian Perjalanan Dinas (5.1.02.04)",
                                    "Total Belanja Perjalanan Dinas")
        
        # AUTO-DISPLAY BKU DATA
        self._display_bku_for_category("Perjalanan Dinas")

    def show_peralatan(self):
        """Show both RKAS and BKU data for Peralatan"""
        if not self.processor.peralatan_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        
        # Display RKAS data
        self.display_standard_results(self.processor.peralatan_items,
                                    "Rincian Peralatan dan Mesin (5.2.02)",
                                    "Total Peralatan")
        
        # AUTO-DISPLAY BKU DATA
        self._display_bku_for_category("Peralatan")

    def show_aset_tetap(self):
        """Show RKAS data for Aset Tetap"""
        if not self.processor.aset_tetap_items:
            messagebox.showwarning("Peringatan", "Data dalam kategori tersebut tidak ada atau file belum diupload!")
            return
        self.display_standard_results(self.processor.aset_tetap_items,
                                    "Rincian Aset Tetap Lainnya (5.2.04 & 5.2.05)",
                                    "Total Aset Tetap Lainnya")
        
        # Clear BKU section for non-supported categories
        self._display_bku_for_category("Aset Tetap")

    def show_ringkasan(self):
        """Show summary data - FIXED VERSION dengan debug"""
        if not self.processor.excel_data and self.processor.total_penerimaan == 0:
            messagebox.showwarning("Peringatan", "File belum diupload!")
            return
        
        # Display RKAS summary
        columns = ('Kategori', 'Jumlah (Rp)')
        self.create_standard_table("Ringkasan Anggaran", columns)
        
        # Get summary data from processor
        summary_data = self.processor.get_summary_data()
        
        # Data ringkasan dengan background color info
        ringkasan_data = [
            ("PAGU TAHUN 2025", self.processor.total_penerimaan, True),
            ("BELANJA OPERASI", summary_data['total_belanja_operasi'], True),
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
        
        # Add school name label using consistent layout
        sekolah_label = tk.Label(self.summary_frame, text=f"SEKOLAH {self.processor.nama_sekolah}", 
                                font=('Arial', 10, 'bold'), bg='#ecf0f1', fg='#2c3e50')
        sekolah_label.pack(anchor='w', pady=5)
        
        # DEBUG: Print BKU info
        print(f"Debug: Checking BKU data availability...")
        print(f"Debug: hasattr bku_data_available: {hasattr(self.processor, 'bku_data_available')}")
        if hasattr(self.processor, 'bku_data_available'):
            print(f"Debug: bku_data_available value: {self.processor.bku_data_available}")
        
        # AUTO-DISPLAY BKU SUMMARY untuk triwulan yang dipilih
        if hasattr(self.processor, 'bku_data_available') and self.processor.bku_data_available:
            selected_triwulan = self.selected_triwulan.get()
            print(f"Debug: Displaying BKU summary for {selected_triwulan}")
            
            # DEBUG: Cek summary data sebelum display
            summary_data_bku = self.processor.get_bku_summary_data_by_triwulan(selected_triwulan)
            print(f"Debug: BKU summary data: {summary_data_bku}")
            
            self._display_bku_summary_for_triwulan(selected_triwulan)
        else:
            print("Debug: BKU data not available, showing placeholder")
            # Clear BKU section jika tidak ada data
            self._clear_bku_for_non_supported()
            
    def _clear_bku_for_non_supported(self):
            """Clear BKU section and show placeholder for non-supported categories"""
            # Clear current BKU content
            for widget in self.bku_frame.winfo_children():
                widget.destroy()
            
            # Reset BKU canvas scroll position to top
            self.bku_canvas.yview_moveto(0)
            
            # Show placeholder
            self._create_bku_placeholder()
            
            # Update canvas scroll region
            self.bku_frame.update_idletasks()
            self.bku_canvas.configure(scrollregion=self.bku_canvas.bbox("all"))

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
        if hasattr(self, 'canvas_window'):
            self.button_canvas.itemconfig(self.canvas_window, height=canvas_height)
    
    def _on_frame_configure(self, event):
        """Handle frame configure event"""
        # Update scroll region when frame changes
        self.button_canvas.configure(scrollregion=self.button_canvas.bbox("all"))


    # Tambahkan method-method yang hilang ini ke dalam class RKASPage

    def _create_navigation_buttons(self):
        """Create navigation buttons for different categories"""
        # Define button configurations
        button_configs = [
            ("Belanja Persediaan", self.show_belanja_persediaan, '#3498db'),
            ("Belanja Jasa", self.show_belanja_jasa, '#e67e22'),
            ("Pemeliharaan", self.show_belanja_pemeliharaan, '#9b59b6'),
            ("Perjalanan Dinas", self.show_belanja_perjalanan, '#1abc9c'),
            ("Peralatan", self.show_peralatan, '#34495e'),
            ("Aset Tetap", self.show_aset_tetap, '#e74c3c'),
            ("Ringkasan", self.show_ringkasan, '#27ae60')
        ]
        
        # Clear existing buttons
        self.tab_buttons.clear()
        
        # Create buttons
        for text, command, color in button_configs:
            btn = tk.Button(
                self.scrollable_button_frame,
                text=text,
                command=lambda cmd=command, name=text: self._handle_button_click(cmd, name),
                bg=color,
                fg='white',
                font=('Arial', 11, 'bold'),
                padx=20,
                pady=8,
                cursor='hand2',
                relief='raised',
                bd=2
            )
            btn.pack(side='left', padx=5, pady=5)
            
            # Store button reference with original color
            self.tab_buttons[text] = (btn, color)
        
        # Update scroll region
        self.scrollable_button_frame.update_idletasks()
        self.button_canvas.configure(scrollregion=self.button_canvas.bbox("all"))

    def _handle_button_click(self, command, tab_name):
        """Handle button click with highlighting"""
        # Reset all button colors
        for name, (btn, orig_color) in self.tab_buttons.items():
            btn.config(bg=orig_color)
        
        # Highlight active button
        if tab_name in self.tab_buttons:
            active_btn, _ = self.tab_buttons[tab_name]
            active_btn.config(bg='#2c3e50')  # Dark highlight color
        
        # Set active tab
        self.active_tab = tab_name
        
        # Execute command
        command()

    def _setup_canvas_scrolling(self):
        """Setup canvas scrolling for button section"""
        # Bind frame configure event
        self.scrollable_button_frame.bind('<Configure>', self._on_frame_configure)
        
        # Bind canvas configure event
        self.button_canvas.bind('<Configure>', self._on_canvas_configure)
        
        # Bind mouse wheel events for horizontal scrolling
        self.button_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.button_canvas.bind("<Button-4>", self._on_mousewheel)  # Linux
        self.button_canvas.bind("<Button-5>", self._on_mousewheel)  # Linux





