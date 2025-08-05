"""
Home page for SIKELAR application
Contains the main tool selection interface with modern design
"""

import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont
from .base_page import BasePage

class HomePage(BasePage):
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app)
        self.canvas = None
        self.mousewheel_bound = False
        
    def build_page(self):
        """Build the home page content with modern design"""
        # Create main canvas and scrollbar for home page
        self.canvas = tk.Canvas(self.page_frame, bg='#f5f6fa', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.page_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f5f6fa')
        
        # Pack canvas and scrollbar first
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Configure scrolling AFTER packing
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.update_scroll_region()
        )
        
        # Create window in canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Bind canvas resize
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Header section with blue background
        header_frame = tk.Frame(self.scrollable_frame, bg='#4a69bd', height=180)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        # Header content container
        header_content = tk.Frame(header_frame, bg='#4a69bd')
        header_content.place(relx=0.5, rely=0.5, anchor='center')
        
        title_label = tk.Label(
            header_content,
            text="SIKELAR",
            font=tkFont.Font(family="Segoe UI", size=36, weight="bold"),
            bg='#4a69bd',
            fg='white'
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            header_content,
            text="Sistem Informasi Pengelompokan Anggaran dan Rekening",
            font=tkFont.Font(family="Segoe UI", size=14),
            bg='#4a69bd',
            fg='#e8f2ff'
        )
        subtitle_label.pack(pady=(8, 0))
        
        # Main content area
        content_frame = tk.Frame(self.scrollable_frame, bg='#f5f6fa', padx=60, pady=40)
        content_frame.pack(fill='both', expand=True)
        
        # Section title
        section_title = tk.Label(
            content_frame,
            text="Pilih Menu Aplikasi:",
            font=tkFont.Font(family="Segoe UI", size=18, weight="bold"),
            bg='#f5f6fa',
            fg='#2c3e50',
            anchor='w'
        )
        section_title.pack(fill='x', pady=(0, 25))
        
        # Tools grid container
        tools_frame = tk.Frame(content_frame, bg='#f5f6fa')
        tools_frame.pack(fill='x', pady=(0, 30))
        
        # Create tool cards
        self.create_tool_cards(tools_frame)
        
        # Additional info section
        info_frame = tk.Frame(content_frame, bg='#ffffff', relief='flat', bd=0)
        info_frame.pack(fill='x', pady=(20, 0))
        
        # Add subtle shadow to info frame
        shadow_frame = tk.Frame(content_frame, bg='#e2e8f0', height=2)
        shadow_frame.pack(fill='x', pady=(22, 0))
        shadow_frame.lower()
        
        info_content = tk.Frame(info_frame, bg='#ffffff', padx=25, pady=20)
        info_content.pack(fill='both', expand=True)
        
        info_title = tk.Label(
            info_content,
            text="📋 Petunjuk Penggunaan",
            font=tkFont.Font(family="Segoe UI", size=16, weight="bold"),
            bg='#ffffff',
            fg='#2c3e50',
            anchor='w'
        )
        info_title.pack(fill='x', pady=(0, 10))
        
        info_text = tk.Label(
            info_content,
            text="• Pilih menu 'RKAS dan BKU' untuk melihat perhitungan RKAS dan Realisasi BKU dengan mengupload 1 file dalam bentuk .xlsx (sheet 1: RKAS, sheet 2: BKU)\n• Pilih menu 'Pengesahan' untuk melihat detail persentase dari pagu untuk buku, sarana prasarana, dan honor",
            font=tkFont.Font(family="Segoe UI", size=11),
            bg='#ffffff',
            fg='#6c757d',
            anchor='w',
            justify='left'
        )
        info_text.pack(fill='x')
        
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
            text="© 2025 SIKELAR - Sistem Informasi Pengelompokan Anggaran dan Rekening",
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
        
    def on_show(self):
        """Called when home page is shown"""
        # Bind mousewheel to canvas
        self.bind_mousewheel()
        
        # Force update of canvas size and scroll region after everything is created
        self.main_app.root.after(10, self.force_update_canvas)
    
    def on_hide(self):
        """Called when home page is hidden"""
        # Unbind mousewheel before hiding page
        self.unbind_mousewheel()
        
    def force_update_canvas(self):
        """Force update canvas size and scroll region"""
        if self.canvas and self.canvas.winfo_exists():
            # Update canvas
            self.canvas.update_idletasks()
            self.scrollable_frame.update_idletasks()
            
            # Set proper canvas window width
            canvas_width = self.canvas.winfo_width()
            if canvas_width > 1:  # Make sure canvas has been sized
                self.canvas.itemconfig(self.canvas_window, width=canvas_width)
            
            # Update scroll region
            self.update_scroll_region()
    
    def update_scroll_region(self):
        """Update the scroll region of the canvas"""
        if self.canvas and self.canvas.winfo_exists():
            self.canvas.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def create_tool_cards(self, parent):
        """Create the grid of tool cards with modern design"""
        tools = [
            {
                'title': 'RKAS dan BKU',
                'desc': 'Upload file berbentuk excel\nuntuk analisis anggaran',
                'color': '#4dabf7',
                'icon': '📊',
                'bg_gradient': '#e3f2fd'
            },
            {
                'title': 'Pengesahan',
                'desc': 'Ctrl + a halaman detail\npengesahan RKAS',
                'color': '#7c3aed',
                'icon': '✅',
                'bg_gradient': '#f3e5f5'
            }
        ]
        
        # Create row with 2 cards
        for col, tool in enumerate(tools):
            self.create_tool_card(parent, tool, col)
    
    def create_tool_card(self, parent, tool, col):
        """Create a single tool card with modern design"""
        # Configure grid weights for even distribution
        parent.grid_columnconfigure(col, weight=1)
        
        # Card container with modern styling
        card_frame = tk.Frame(
            parent,
            bg='#ffffff',
            relief='flat',
            bd=0,
            padx=2,
            pady=2
        )
        card_frame.grid(row=0, column=col, padx=15, pady=10, sticky='ew')
        
        # Add modern shadow effect
        shadow_frame = tk.Frame(
            parent,
            bg='#e2e8f0',
            height=3
        )
        shadow_frame.grid(row=0, column=col, padx=17, pady=13, sticky='ew')
        shadow_frame.lower()
        
        # Card content frame
        content_frame = tk.Frame(card_frame, bg='#ffffff', padx=25, pady=25)
        content_frame.pack(fill='both', expand=True)
        
        # Icon and arrow section
        top_section = tk.Frame(content_frame, bg='#ffffff')
        top_section.pack(fill='x', pady=(0, 20))
        
        # Modern icon design
        icon_container = tk.Frame(top_section, bg='#ffffff')
        icon_container.pack(side='left')
        
        # Icon background with gradient effect
        icon_bg = tk.Frame(
            icon_container,
            bg=tool['color'],
            width=60,
            height=60
        )
        icon_bg.pack()
        icon_bg.pack_propagate(False)
        
        # Icon emoji/symbol
        icon_label = tk.Label(
            icon_bg,
            text=tool['icon'],
            font=tkFont.Font(family="Segoe UI Emoji", size=24),
            bg=tool['color'],
            fg='white'
        )
        icon_label.place(relx=0.5, rely=0.5, anchor='center')
        
        # Modern arrow
        arrow_container = tk.Frame(top_section, bg='#ffffff')
        arrow_container.pack(side='right')
        
        arrow_bg = tk.Frame(
            arrow_container,
            bg='#f8f9fa',
            width=40,
            height=40
        )
        arrow_bg.pack()
        arrow_bg.pack_propagate(False)
        
        arrow_label = tk.Label(
            arrow_bg,
            text="→",
            font=tkFont.Font(family="Segoe UI", size=18, weight="bold"),
            bg='#f8f9fa',
            fg='#6c757d'
        )
        arrow_label.place(relx=0.5, rely=0.5, anchor='center')
        
        # Title with modern typography
        title_label = tk.Label(
            content_frame,
            text=tool['title'],
            font=tkFont.Font(family="Segoe UI", size=18, weight="bold"),
            bg='#ffffff',
            fg='#2c3e50',
            anchor='w'
        )
        title_label.pack(fill='x', pady=(0, 8))
        
        # Description with better spacing
        desc_label = tk.Label(
            content_frame,
            text=tool['desc'],
            font=tkFont.Font(family="Segoe UI", size=11),
            bg='#ffffff',
            fg='#6c757d',
            anchor='w',
            justify='left'
        )
        desc_label.pack(fill='x')
        
        # Add subtle accent line at bottom
        accent_line = tk.Frame(content_frame, bg=tool['color'], height=3)
        accent_line.pack(fill='x', pady=(15, 0))
        
        # Make card clickable with hover effects
        self.make_clickable_modern(card_frame, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(content_frame, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(top_section, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(icon_container, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(icon_bg, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(icon_label, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(title_label, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(desc_label, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(arrow_container, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(arrow_bg, lambda t=tool['title']: self.main_app.show_tool_page(t))
        self.make_clickable_modern(arrow_label, lambda t=tool['title']: self.main_app.show_tool_page(t))
    
    def make_clickable_modern(self, widget, callback):
        """Make widget clickable with modern hover effects"""
        original_bg = widget.cget('bg')
        
        def on_enter(e):
            widget.configure(cursor='hand2')
            # Subtle hover effect
            if original_bg == '#ffffff':
                widget.configure(bg='#f8f9fa')
            
        def on_leave(e):
            widget.configure(cursor='')
            widget.configure(bg=original_bg)
            
        def on_click(e):
            # Brief click animation
            widget.configure(bg='#e9ecef')
            widget.after(100, lambda: widget.configure(bg=original_bg))
            callback()
            
        widget.bind('<Button-1>', on_click)
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)
    
    def on_canvas_configure(self, event):
        """Handle canvas resize events"""
        if self.canvas and self.canvas.winfo_exists():
            canvas_width = event.width
            if canvas_width > 1:  # Make sure we have a valid width
                self.canvas.itemconfig(self.canvas_window, width=canvas_width)
                # Update scroll region when canvas is resized
                self.main_app.root.after_idle(self.update_scroll_region)
    
    def bind_mousewheel(self):
        """Bind mousewheel events for scrolling"""
        if self.mousewheel_bound or not self.canvas:
            return
            
        def on_mousewheel(event):
            if self.canvas and self.canvas.winfo_exists():
                # Check if canvas is visible and has scrollable content
                try:
                    bbox = self.canvas.bbox("all")
                    if bbox:
                        canvas_height = self.canvas.winfo_height()
                        content_height = bbox[3] - bbox[1]
                        
                        # Only scroll if content is larger than canvas
                        if content_height > canvas_height:
                            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                except tk.TclError:
                    # Canvas might be destroyed, ignore
                    pass
        
        # Bind mousewheel to canvas and its children
        if self.canvas and self.canvas.winfo_exists():
            self.canvas.bind("<MouseWheel>", on_mousewheel)
            
            # Also bind to the scrollable_frame to catch events when mouse is over content
            if hasattr(self, 'scrollable_frame') and self.scrollable_frame.winfo_exists():
                self.scrollable_frame.bind("<MouseWheel>", on_mousewheel)
                
                # Bind to all child widgets recursively
                self.bind_mousewheel_recursive(self.scrollable_frame, on_mousewheel)
            
            self.mousewheel_bound = True
    
    def bind_mousewheel_recursive(self, widget, callback):
        """Recursively bind mousewheel to all child widgets"""
        try:
            widget.bind("<MouseWheel>", callback, add='+')
            for child in widget.winfo_children():
                self.bind_mousewheel_recursive(child, callback)
        except tk.TclError:
            # Widget might be destroyed, ignore
            pass
    
    def unbind_mousewheel(self):
        """Unbind mousewheel events"""
        if not self.mousewheel_bound:
            return
            
        try:
            if self.canvas and self.canvas.winfo_exists():
                self.canvas.unbind("<MouseWheel>")
                
            if hasattr(self, 'scrollable_frame') and self.scrollable_frame.winfo_exists():
                self.scrollable_frame.unbind("<MouseWheel>")
                self.unbind_mousewheel_recursive(self.scrollable_frame)
                
        except tk.TclError:
            # Widgets might be destroyed, ignore
            pass
        finally:
            self.mousewheel_bound = False
    
    def unbind_mousewheel_recursive(self, widget):
        """Recursively unbind mousewheel from all child widgets"""
        try:
            widget.unbind("<MouseWheel>")
            for child in widget.winfo_children():
                self.unbind_mousewheel_recursive(child)
        except tk.TclError:
            # Widget might be destroyed, ignore
            pass