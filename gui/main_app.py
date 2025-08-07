"""
Main application class for SIKELAR - FIXED VERSION
Handles navigation between pages and common functionality
"""

import tkinter as tk
import tkinter.font as tkFont
from .pages.home_page import HomePage
from .pages.pengesahan_page import PengesahanPage  
from .pages.rkas_page import RKASPage
from backend.processor import BOSDataProcessor  # PERBAIKAN: Import yang benar

class SikelarMainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SIKELAR")
        self.root.geometry("1200x800")  # Lebih besar untuk split view
        self.root.configure(bg='#f8f9fa')
        
        # Initialize data processor (shared across pages)
        self.data_processor = BOSDataProcessor()  # PERBAIKAN: Gunakan BOSDataProcessor
        
        # Configure common styles
        self.setup_styles()
        
        # Create main container frame that will hold all pages
        self.main_container = tk.Frame(self.root, bg='#f8f9fa')
        self.main_container.pack(fill='both', expand=True)
        
        # Current page tracker
        self.current_page = None
        
        # Initialize pages dictionary
        self.pages = {}
        
        # Show home page immediately after widget creation
        self.root.after(1, self.show_home_page)
        
    def setup_styles(self):
        """Setup common styles for the application"""
        self.title_font = tkFont.Font(family="Segoe UI", size=24, weight="bold")
        self.subtitle_font = tkFont.Font(family="Segoe UI", size=12)
        self.card_title_font = tkFont.Font(family="Segoe UI", size=14, weight="bold")
        self.card_desc_font = tkFont.Font(family="Segoe UI", size=10)
        self.page_title_font = tkFont.Font(family="Segoe UI", size=20, weight="bold")
        
    def get_or_create_page(self, page_name):
        """Get existing page or create new one (lazy loading)"""
        if page_name not in self.pages:
            if page_name == 'home':
                self.pages['home'] = HomePage(self.main_container, self)
            elif page_name == 'pengesahan':
                self.pages['pengesahan'] = PengesahanPage(self.main_container, self)
            elif page_name == 'rkas':
                self.pages['rkas'] = RKASPage(self.main_container, self)
        
        return self.pages[page_name]
        
    def clear_page(self):
        """Clear current page content"""
        if self.current_page:
            self.current_page.hide()
        
    def show_home_page(self):
        """Show home page"""
        self.clear_page()
        self.current_page = self.get_or_create_page('home')
        self.current_page.show()
        
    def show_pengesahan_page(self):
        """Show pengesahan page"""
        self.clear_page()
        self.current_page = self.get_or_create_page('pengesahan')
        self.current_page.show()
        
    def show_rkas_page(self):
        """Show RKAS page"""
        self.clear_page()
        self.current_page = self.get_or_create_page('rkas')
        self.current_page.show()
        
    def show_tool_page(self, tool_name):
        """Show the appropriate page based on tool name"""
        if tool_name == "Pengesahan":
            self.show_pengesahan_page()
        elif tool_name == "Rekon":
            self.show_rkas_page()
        else:
            # For any other tools, show blank page
            self.show_blank_tool_page(tool_name)
    
    def show_blank_tool_page(self, tool_name):
        """Show a blank page for tools that don't have dedicated pages yet"""
        self.clear_page()
        
        # Create temporary page frame
        temp_page = tk.Frame(self.main_container, bg='#f8f9fa')
        temp_page.pack(fill='both', expand=True)
        self.current_page = temp_page
        
        # Header with back button
        header_frame = tk.Frame(temp_page, bg='#f8f9fa', pady=20)
        header_frame.pack(fill='x', padx=20)
        
        # Back button
        back_btn = tk.Button(
            header_frame,
            text="← Back to Home",
            font=tkFont.Font(family="Segoe UI", size=12),
            bg='#6c757d',
            fg='white',
            padx=20,
            pady=8,
            border=0,
            cursor='hand2',
            command=self.show_home_page
        )
        back_btn.pack(side='left')
        
        # Tool page title
        title_frame = tk.Frame(temp_page, bg='#f8f9fa', pady=40)
        title_frame.pack(fill='x')
        
        title_label = tk.Label(
            title_frame,
            text=f"{tool_name}",
            font=self.page_title_font,
            bg='#f8f9fa',
            fg='#2c3e50'
        )
        title_label.pack()
        
        # Content area (blank for now)
        content_frame = tk.Frame(temp_page, bg='white', relief='solid', bd=1)
        content_frame.pack(fill='both', expand=True, padx=50, pady=20)
        
        # Placeholder text
        placeholder_label = tk.Label(
            content_frame,
            text=f"This is the {tool_name} page.\nFeatures will be added here.",
            font=tkFont.Font(family="Segoe UI", size=14),
            bg='white',
            fg='#6c757d',
            justify='center'
        )
        placeholder_label.pack(expand=True)