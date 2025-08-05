"""
Pengesahan page for SIKELAR application
Enhanced UI with modern design, full scrolling support, and improved symmetrical layout
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox
import tkinter.font as tkFont
from .base_page import BasePage
from backend.utils import FormatUtils

class PengesahanPage(BasePage):
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app)
        self.active_button = None
        self.category_buttons = {}
        
    def build_page(self):
        """Build the pengesahan page content with enhanced modern UI and improved symmetrical layout"""
        self.page_frame.configure(bg='#f8f9fa')
        
        # SCROLLABLE CONTENT AREA - Main container yang bisa di-scroll
        # Create canvas and scrollbar for scrollable content
        self.canvas = tk.Canvas(self.page_frame, bg='#f8f9fa', highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self.page_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f8f9fa')
        
        # Configure scrollable frame
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window in canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Configure canvas scroll
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel to canvas for full area scrolling
        self.bind_mousewheel()
        
        # Configure canvas window to resize with canvas
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Enhanced Header matching SIKELAR homepage colors (unchanged)
        header_frame = tk.Frame(self.scrollable_frame, bg='#4a69bd', height=100)
        header_frame.pack(fill='x', side='top')
        header_frame.pack_propagate(False)
        
        # Back button positioned absolutely at left with improved styling
        back_btn = tk.Button(
            header_frame,
            text="🏠 Home",
            font=tkFont.Font(family="Segoe UI", size=11, weight="bold"),
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
        
        # Add subtle hover effects for back button
        back_btn.bind("<Enter>", lambda e: self.on_hover(e, '#2c3393'))
        back_btn.bind("<Leave>", lambda e: self.on_leave(e, '#3742a6'))
        
        # Centered title container - positioned in center regardless of back button
        title_container = tk.Frame(header_frame, bg='#4a69bd')
        title_container.place(relx=0.5, rely=0.5, anchor='center')
        
        # Main title - perfectly centered (unchanged)
        header_label = tk.Label(title_container, 
                               text="Sistem Pemeriksaan Pengesahan RKAS",
                               font=("Segoe UI", 20, "bold"), 
                               bg='#4a69bd', 
                               fg='white')
        header_label.pack()
        
        # Subtitle for better visual hierarchy (unchanged)
        subtitle_label = tk.Label(title_container,
                                 text="Analisis dan Kategorisasi Anggaran Sekolah",
                                 font=("Segoe UI", 11),
                                 bg='#4a69bd',
                                 fg='#c8d4f0')
        subtitle_label.pack()
        
        # Main container with improved spacing - INSIDE SCROLLABLE FRAME
        main_container = tk.Frame(self.scrollable_frame, bg='#f8f9fa')
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=25)
        
        # Enhanced Upload section with card-like design and improved layout
        upload_card = tk.Frame(main_container, bg='white', relief='flat', bd=0)
        upload_card.pack(fill=tk.X, pady=(0, 25))
        
        # Add subtle shadow effect
        shadow_frame = tk.Frame(main_container, bg='#e9ecef', height=2)
        shadow_frame.pack(fill=tk.X, pady=(0, 23))
        
        upload_inner = tk.Frame(upload_card, bg='white')
        upload_inner.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        upload_label = tk.Label(upload_inner, 
                               text="📋 Paste Detail Pengesahan RKAS di Sini:", 
                               font=("Segoe UI", 12, "bold"), 
                               bg='white',
                               fg='#2c3e50')
        upload_label.pack(anchor="w", pady=(0, 10))
        
        # Modified layout: Input area and buttons side by side with better proportions
        input_button_frame = tk.Frame(upload_inner, bg='white')
        input_button_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input area (left side) - improved width ratio
        input_frame = tk.Frame(input_button_frame, bg='#f8f9fa', relief='flat', bd=1)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 20))
        
        self.input_text = scrolledtext.ScrolledText(input_frame, 
                                                   height=8, 
                                                   wrap=tk.WORD,
                                                   font=("Segoe UI", 10), 
                                                   bg='white',
                                                   relief='flat', 
                                                   bd=0,
                                                   selectbackground='#3498db',
                                                   selectforeground='white',
                                                   insertbackground='#2c3e50')
        self.input_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # Button area (right side) with improved spacing and symmetry
        button_frame = tk.Frame(input_button_frame, bg='white')
        button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 5))
        
        # Process button with improved styling and consistent size
        self.process_button = tk.Button(button_frame, 
                                       text="🚀 Proses Data", 
                                       command=self.process_data,
                                       bg='#28a745', 
                                       fg='white', 
                                       font=("Segoe UI", 11, "bold"),
                                       relief='flat', 
                                       padx=20, 
                                       pady=14,
                                       cursor='hand2',
                                       bd=0,
                                       width=16,
                                       height=2)
        self.process_button.pack(pady=(0, 12), fill=tk.X)
        
        # Enhanced hover effects
        self.process_button.bind("<Enter>", lambda e: self.on_hover(e, '#218838'))
        self.process_button.bind("<Leave>", lambda e: self.on_leave(e, '#28a745'))
        
        # Clear button with consistent styling
        self.clear_button = tk.Button(button_frame, 
                                     text="🗑️ Bersihkan", 
                                     command=self.clear_data,
                                     bg='#dc3545', 
                                     fg='white', 
                                     font=("Segoe UI", 11, "bold"),
                                     relief='flat', 
                                     padx=20, 
                                     pady=14,
                                     cursor='hand2',
                                     bd=0,
                                     width=16,
                                     height=2)
        self.clear_button.pack(fill=tk.X)
        
        self.clear_button.bind("<Enter>", lambda e: self.on_hover(e, '#c82333'))
        self.clear_button.bind("<Leave>", lambda e: self.on_leave(e, '#dc3545'))
        
        # Enhanced Category buttons with improved symmetry and consistent sizing
        category_container = tk.Frame(main_container, bg='#f8f9fa')
        category_container.pack(fill=tk.X, pady=(15, 25))
        
        category_label = tk.Label(category_container, 
                                 text="📊 Kategori Anggaran:",
                                 font=("Segoe UI", 12, "bold"),
                                 bg='#f8f9fa',
                                 fg='#2c3e50')
        category_label.pack(anchor="w", pady=(0, 15))
        
        # Center the category buttons for better symmetry
        category_frame = tk.Frame(category_container, bg='#f8f9fa')
        category_frame.pack(anchor='center')
        
        # BUKU button with improved design and consistent sizing
        self.btn_buku = tk.Button(category_frame, 
                                 text="📚 BUKU", 
                                 command=self.show_alokasi_buku, 
                                 state=tk.DISABLED,
                                 bg='#007bff', 
                                 fg='white', 
                                 font=("Segoe UI", 10, "bold"),
                                 relief='flat', 
                                 padx=30, 
                                 pady=18, 
                                 cursor='hand2',
                                 bd=0,
                                 width=14,
                                 height=2)
        self.btn_buku.pack(side=tk.LEFT, padx=8, fill=tk.Y)
        
        # SARANA DAN PRASARANA button with consistent styling
        self.btn_sarana = tk.Button(category_frame, 
                                   text="🏢 SARANA & PRASARANA", 
                                   command=self.show_alokasi_sarana, 
                                   state=tk.DISABLED,
                                   bg='#6f42c1', 
                                   fg='white', 
                                   font=("Segoe UI", 10, "bold"),
                                   relief='flat', 
                                   padx=30, 
                                   pady=18, 
                                   cursor='hand2',
                                   bd=0,
                                   width=20,
                                   height=2)
        self.btn_sarana.pack(side=tk.LEFT, padx=8, fill=tk.Y)
        
        # HONOR button with consistent styling
        self.btn_honor = tk.Button(category_frame, 
                                  text="💰 HONOR", 
                                  command=self.show_alokasi_honor, 
                                  state=tk.DISABLED,
                                  bg='#fd7e14', 
                                  fg='white', 
                                  font=("Segoe UI", 10, "bold"),
                                  relief='flat', 
                                  padx=30, 
                                  pady=18, 
                                  cursor='hand2',
                                  bd=0,
                                  width=14,
                                  height=2)
        self.btn_honor.pack(side=tk.LEFT, padx=8, fill=tk.Y)
        
        # Store category buttons for easy access
        self.category_buttons = {
            'buku': self.btn_buku,
            'sarana': self.btn_sarana,
            'honor': self.btn_honor
        }
        
        # Setup enhanced hover effects
        self.setup_category_button_hover()
        
        # Enhanced Results section with card design - header color unchanged
        results_card = tk.Frame(main_container, bg='white', relief='flat', bd=0)
        results_card.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        # Results header with matching gradient (unchanged)
        results_header = tk.Frame(results_card, bg='#4a69bd', height=50)
        results_header.pack(fill=tk.X)
        results_header.pack_propagate(False)
        
        results_label = tk.Label(results_header, 
                                text="📋 Hasil Analisis:", 
                                font=("Segoe UI", 12, "bold"), 
                                bg='#4a69bd', 
                                fg='white')
        results_label.pack(side=tk.LEFT, padx=20, pady=15)
        
        # Results content with enhanced styling - consistent with input area
        self.output_text = scrolledtext.ScrolledText(results_card, 
                                                    height=18,  # Slightly reduced for better proportions
                                                    wrap=tk.WORD, 
                                                    font=("Consolas", 10),
                                                    bg='#fdfdfd', 
                                                    relief='flat', 
                                                    bd=0,
                                                    selectbackground='#007bff',
                                                    selectforeground='white',
                                                    insertbackground='#2c3e50')
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

         # Footer section
        self.create_footer()
    
    def create_footer(self):
        """Create footer section with modern design matching the theme - FIXED VERSION"""
        # PERBAIKAN: Pastikan footer dibuat dengan benar dan semua elemen ditampilkan
        
        # Footer container dengan background gelap
        footer_frame = tk.Frame(self.scrollable_frame, bg='#2c3e50')
        footer_frame.pack(fill='x', pady=(40, 0))
        
        # PERBAIKAN: Jangan gunakan pack_propagate(False) yang bisa menyembunyikan konten
        # footer_frame.pack_propagate(False)
        
        # Footer content container dengan padding yang cukup
        footer_content = tk.Frame(footer_frame, bg='#2c3e50')
        footer_content.pack(fill='both', expand=True, padx=60, pady=25)
        
        # Top section dengan app info dan links
        footer_top = tk.Frame(footer_content, bg='#2c3e50')
        footer_top.pack(fill='x', pady=(0, 15))
        
        # Left side - App information
        footer_left = tk.Frame(footer_top, bg='#2c3e50')
        footer_left.pack(side='left', fill='both', expand=True)
        
        app_name_footer = tk.Label(
            footer_left,
            text="SIKELAR",
            font=tkFont.Font(family="Segoe UI", size=16, weight="bold"),
            bg='#2c3e50',
            fg='#ffffff'
        )
        app_name_footer.pack(anchor='w')
        
        app_desc_footer = tk.Label(
            footer_left,
            text="Sistem Informasi Pengelompokan Anggaran dan Rekening",
            font=tkFont.Font(family="Segoe UI", size=10),
            bg='#2c3e50',
            fg='#bdc3c7'
        )
        app_desc_footer.pack(anchor='w', pady=(2, 0))
        
        # Right side - Quick links
        footer_right = tk.Frame(footer_top, bg='#2c3e50')
        footer_right.pack(side='right')
        
        links_title = tk.Label(
            footer_right,
            text="Menu Cepat",
            font=tkFont.Font(family="Segoe UI", size=12, weight="bold"),
            bg='#2c3e50',
            fg='#ffffff'
        )
        links_title.pack(anchor='e')
        
        # Quick links container
        links_container = tk.Frame(footer_right, bg='#2c3e50')
        links_container.pack(anchor='e', pady=(5, 0))
        
        # RKAS dan BKU link
        rkas_link = tk.Label(
            links_container,
            text="• RKAS dan BKU",
            font=tkFont.Font(family="Segoe UI", size=10),
            bg='#2c3e50',
            fg='#4dabf7',
            cursor='hand2'
        )
        rkas_link.pack(anchor='e')
        
        # Pengesahan link
        pengesahan_link = tk.Label(
            links_container,
            text="• Pengesahan",
            font=tkFont.Font(family="Segoe UI", size=10),
            bg='#2c3e50',
            fg='#7c3aed',
            cursor='hand2'
        )
        pengesahan_link.pack(anchor='e', pady=(2, 0))
        
        # Make links clickable
        rkas_link.bind('<Button-1>', lambda e: self.main_app.show_tool_page('RKAS dan BKU'))
        pengesahan_link.bind('<Button-1>', lambda e: self.main_app.show_tool_page('Pengesahan'))
        
        # Add hover effects to links
        self.add_link_hover_effect(rkas_link, '#4dabf7', '#74c0fc')
        self.add_link_hover_effect(pengesahan_link, '#7c3aed', '#9775fa')
        
        # Bottom section dengan copyright dan version
        footer_bottom = tk.Frame(footer_content, bg='#2c3e50')
        footer_bottom.pack(fill='x')
        
        # Separator line
        separator = tk.Frame(footer_bottom, bg='#34495e', height=1)
        separator.pack(fill='x', pady=(0, 10))
        
        # Bottom content
        footer_bottom_content = tk.Frame(footer_bottom, bg='#2c3e50')
        footer_bottom_content.pack(fill='x')
        
        # Copyright text
        copyright_text = tk.Label(
            footer_bottom_content,
            text="© 2024 SIKELAR - Sistem Informasi Pengelompokan Anggaran dan Rekening",
            font=tkFont.Font(family="Segoe UI", size=9),
            bg='#2c3e50',
            fg='#95a5a6'
        )
        copyright_text.pack(side='left')
        
        # Version info
        version_text = tk.Label(
            footer_bottom_content,
            text="Versi 1.0",
            font=tkFont.Font(family="Segoe UI", size=9),
            bg='#2c3e50',
            fg='#95a5a6'
        )
        version_text.pack(side='right')
        
        # PERBAIKAN: Pastikan footer memiliki tinggi minimum yang cukup
        footer_frame.configure(height=120)
        footer_frame.update_idletasks()
    
    def add_link_hover_effect(self, widget, normal_color, hover_color):
        """Add hover effect to footer links"""
        def on_enter(e):
            widget.configure(fg=hover_color)
            
        def on_leave(e):
            widget.configure(fg=normal_color)
            
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)
    
    
    def bind_mousewheel(self):
        """Bind mouse wheel events to the entire page for full scrolling support"""
        def _on_mousewheel(event):
            # Check if the mouse is over the main canvas area (not over ScrolledText widgets)
            widget_under_mouse = self.page_frame.winfo_containing(event.x_root, event.y_root)
            
            # Don't interfere with ScrolledText internal scrolling
            if (widget_under_mouse != self.input_text and 
                widget_under_mouse != self.output_text and
                not self._is_child_of_scrolledtext(widget_under_mouse)):
                
                # Tambahkan batasan scroll ini:
                current_view = self.canvas.canvasy(0)
                if event.delta > 0 and current_view <= 0:
                    # Jika scroll ke atas dan sudah di posisi paling atas, jangan scroll lagi
                    return
                
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            # Bind to the entire page frame and all its children
            self._bind_mousewheel_recursive(self.page_frame, _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            # Unbind from all widgets
            self._unbind_mousewheel_recursive(self.page_frame)
        
        # Bind events to page frame
        self.page_frame.bind('<Enter>', _bind_to_mousewheel)
        self.page_frame.bind('<Leave>', _unbind_from_mousewheel)
        
        # Also bind directly to main widgets for better coverage
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        # Initial binding
        self._bind_mousewheel_recursive(self.page_frame, _on_mousewheel)
    
    def _is_child_of_scrolledtext(self, widget):
        """Check if widget is a child of any ScrolledText widget"""
        if widget is None:
            return False
        
        # Check if it's one of our ScrolledText widgets or their children
        try:
            parent = widget
            while parent:
                if parent == self.input_text or parent == self.output_text:
                    return True
                parent = parent.master
        except:
            pass
        return False
    
    def _bind_mousewheel_recursive(self, widget, callback):
        """Recursively bind mousewheel to widget and all its children"""
        try:
            widget.bind("<MouseWheel>", callback, add='+')
            for child in widget.winfo_children():
                self._bind_mousewheel_recursive(child, callback)
        except:
            pass
    
    def _unbind_mousewheel_recursive(self, widget):
        """Recursively unbind mousewheel from widget and all its children"""
        try:
            widget.unbind("<MouseWheel>")
            for child in widget.winfo_children():
                self._unbind_mousewheel_recursive(child)
        except:
            pass
    
    def on_canvas_configure(self, event):
        """Handle canvas resize to update scrollable frame width"""
        # Update the scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Update the width of the scrollable frame to match canvas width
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)
    
    def scroll_to_top(self):
        """Scroll to top of the content"""
        self.canvas.yview_moveto(0)
    
    def scroll_to_bottom(self):
        """Scroll to bottom of the content"""
        self.canvas.yview_moveto(1)
    
    def setup_category_button_hover(self):
        """Setup enhanced hover effects untuk category buttons"""
        # Enhanced hover untuk btn_buku
        self.btn_buku.bind("<Enter>", self.on_category_hover_enter)
        self.btn_buku.bind("<Leave>", self.on_category_hover_leave)
        
        # Enhanced hover untuk btn_sarana
        self.btn_sarana.bind("<Enter>", self.on_category_hover_enter)
        self.btn_sarana.bind("<Leave>", self.on_category_hover_leave)
        
        # Enhanced hover untuk btn_honor
        self.btn_honor.bind("<Enter>", self.on_category_hover_enter)
        self.btn_honor.bind("<Leave>", self.on_category_hover_leave)
    
    def set_active_button(self, button_name):
        """Set button aktif dengan enhanced visual state"""
        # Reset semua button ke state normal dengan warna yang sudah diperbaiki
        button_colors = {
            'buku': {'normal': '#007bff', 'active': '#0056b3'},
            'sarana': {'normal': '#6f42c1', 'active': '#5a2d8f'},
            'honor': {'normal': '#fd7e14', 'active': '#dc6003'}
        }
        
        for name, button in self.category_buttons.items():
            if button['state'] == 'normal':
                if name == button_name:
                    # Set button aktif dengan warna lebih gelap dan border
                    button.config(bg=button_colors[name]['active'], 
                                 relief='solid', 
                                 bd=2,
                                 borderwidth=2)
                else:
                    # Reset button lain ke warna normal
                    button.config(bg=button_colors[name]['normal'], 
                                 relief='flat', 
                                 bd=0)
        
        self.active_button = button_name
    
    def get_button_colors(self, button):
        """Get enhanced color scheme untuk button berdasarkan status aktif"""
        button_name = None
        for name, btn in self.category_buttons.items():
            if btn == button:
                button_name = name
                break
        
        color_schemes = {
            'buku': {
                'normal': {'normal': '#007bff', 'hover': '#0056b3'},
                'active': {'normal': '#0056b3', 'hover': '#004085'}
            },
            'sarana': {
                'normal': {'normal': '#6f42c1', 'hover': '#5a2d8f'},
                'active': {'normal': '#5a2d8f', 'hover': '#4a2470'}
            },
            'honor': {
                'normal': {'normal': '#fd7e14', 'hover': '#dc6003'},
                'active': {'normal': '#dc6003', 'hover': '#bd5002'}
            }
        }
        
        if button_name == self.active_button:
            return color_schemes[button_name]['active']
        else:
            return color_schemes[button_name]['normal']
    
    def on_hover(self, event, hover_color):
        """Enhanced event handler saat mouse masuk ke button"""
        event.widget.config(bg=hover_color)
        # Add subtle visual feedback
        event.widget.config(relief='raised', bd=1)
    
    def on_leave(self, event, original_color):
        """Enhanced event handler saat mouse keluar dari button"""
        event.widget.config(bg=original_color)
        event.widget.config(relief='flat', bd=0)
    
    def on_category_hover_enter(self, event):
        """Enhanced event handler saat mouse masuk ke category button"""
        if event.widget['state'] == 'normal':
            colors = self.get_button_colors(event.widget)
            event.widget.config(bg=colors['hover'])
            # Maintain active button border if it's the active one
            if event.widget == self.category_buttons.get(self.active_button):
                event.widget.config(relief='solid', bd=2)
    
    def on_category_hover_leave(self, event):
        """Enhanced event handler saat mouse keluar dari category button"""
        if event.widget['state'] == 'normal':
            colors = self.get_button_colors(event.widget)
            event.widget.config(bg=colors['normal'])
            # Maintain active button styling
            if event.widget == self.category_buttons.get(self.active_button):
                event.widget.config(relief='solid', bd=2)
            else:
                event.widget.config(relief='flat', bd=0)
    
    def clear_data(self):
        """Membersihkan semua data input dan output dengan enhanced feedback"""
        self.input_text.delete("1.0", tk.END)
        self.output_text.delete("1.0", tk.END)
        
        self.main_app.data_processor.clear_data()
        self.active_button = None
        
        # Reset button states dengan warna yang sudah diperbaiki
        button_colors = {
            'buku': '#007bff',
            'sarana': '#6f42c1', 
            'honor': '#fd7e14'
        }
        
        for name, button in self.category_buttons.items():
            button.config(state=tk.DISABLED, 
                         bg=button_colors[name], 
                         relief='flat', 
                         bd=0)
        
        # Scroll to top after clearing
        self.scroll_to_top()
        
        messagebox.showinfo("✅ Berhasil", "Data berhasil dibersihkan!")
    
    def process_data(self):
        """Memproses data input dengan enhanced user feedback"""
        raw_input = self.input_text.get("1.0", tk.END).strip()
        
        if not raw_input:
            messagebox.showwarning("⚠️ Peringatan", "Silakan masukkan data terlebih dahulu!")
            return
        
        try:
            self.active_button = None
            
            result = self.main_app.data_processor.process_data(raw_input)
            
            # Enable semua button dengan warna yang sudah diperbaiki
            button_colors = {
                'buku': '#007bff',
                'sarana': '#6f42c1',
                'honor': '#fd7e14'
            }
            
            for name, button in self.category_buttons.items():
                button.config(state=tk.NORMAL, 
                             bg=button_colors[name], 
                             relief='flat', 
                             bd=0)
            
            self.show_summary()
            
            # Scroll to results section after processing
            self.canvas.update_idletasks()  # Ensure canvas is updated
            self.canvas.yview_moveto(0.6)  # Scroll to show results
            
            messagebox.showinfo("✅ Berhasil", 
                              f"Data berhasil diproses!\n\n"
                              f"🏫 Sekolah: {result['school_name']}\n"
                              f"📊 Total item: {result['total_items']}\n"
                              f"💰 Total Anggaran: {FormatUtils.format_currency(result['total_budget'])}")
            
        except Exception as e:
            messagebox.showerror("❌ Error", f"Terjadi kesalahan saat memproses data:\n{str(e)}")
    
    def show_summary(self):
        """Menampilkan ringkasan data dengan enhanced formatting"""
        self.active_button = None
        button_colors = {
            'buku': '#007bff',
            'sarana': '#6f42c1',
            'honor': '#fd7e14'
        }
        
        for name, button in self.category_buttons.items():
            if button['state'] == 'normal':
                button.config(bg=button_colors[name], relief='flat', bd=0)
        
        output = "═" * 120 + "\n"
        output += f"{'🎯 RINGKASAN DATA YANG DIPROSES':^120}\n"
        output += "═" * 120 + "\n\n"
        output += f"🏫 Nama Sekolah    : {self.main_app.data_processor.school_name}\n"
        output += f"📊 Total Item      : {len(self.main_app.data_processor.processed_data)} item\n"
        output += f"💰 Total Anggaran  : {FormatUtils.format_currency(self.main_app.data_processor.total_budget)}\n\n"
        
        output += "📋 KODE YANG BERHASIL DITEMUKAN:\n"
        output += "─" * 120 + "\n"
        for kode in sorted(self.main_app.data_processor.processed_data.keys()):
            output += f"• {kode:<12} : {self.main_app.data_processor.processed_data[kode]['uraian']}\n"
        
        output += "\n" + "═" * 120 + "\n"
        output += "💡 PETUNJUK: Silakan pilih kategori yang ingin ditampilkan menggunakan tombol di atas.\n"
        output += "═" * 120
        
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", output)
    
    def show_alokasi_buku(self):
        """Menampilkan alokasi buku dengan enhanced display"""
        self.set_active_button('buku')
        
        found_codes = self.main_app.data_processor.get_buku_data()
        
        output = FormatUtils.create_table_display("📚 BUKU (05.02)", 
                                         self.main_app.data_processor.processed_data, 
                                         found_codes,
                                         self.main_app.data_processor.total_budget,
                                         self.main_app.data_processor.school_name)
        
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", output)
        
        # Auto-scroll to show the results
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0.8)
    
    def show_alokasi_sarana(self):
        """Menampilkan alokasi sarana & prasarana dengan enhanced display"""
        self.set_active_button('sarana')
        
        found_codes = self.main_app.data_processor.get_sarana_data()
        
        output = FormatUtils.create_table_display("🏢 SARANA DAN PRASARANA (05.08)", 
                                         self.main_app.data_processor.processed_data, 
                                         found_codes,
                                         self.main_app.data_processor.total_budget,
                                         self.main_app.data_processor.school_name)
        
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", output)
        
        # Auto-scroll to show the results
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0.8)
    
    def show_alokasi_honor(self):
        """Menampilkan alokasi honor dengan enhanced display"""
        self.set_active_button('honor')
        
        found_codes = self.main_app.data_processor.get_honor_data()
        
        output = FormatUtils.create_table_display("💰 HONOR (07.12)", 
                                         self.main_app.data_processor.processed_data, 
                                         found_codes,
                                         self.main_app.data_processor.total_budget,
                                         self.main_app.data_processor.school_name)
        
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", output)
        
        # Auto-scroll to show the results
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0.8)