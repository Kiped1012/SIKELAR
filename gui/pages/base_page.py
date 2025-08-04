"""
Base page class for SIKELAR application
Contains common functionality for all pages
"""

import tkinter as tk

class BasePage:
    def __init__(self, parent, main_app):
        self.parent = parent
        self.main_app = main_app
        self.page_frame = None
        self.is_visible = False
        
    def create_page(self):
        """Create the page frame - to be implemented by subclasses"""
        if not self.page_frame:
            self.page_frame = tk.Frame(self.parent, bg='#f8f9fa')
            self.build_page()
    
    def build_page(self):
        """Build the page content - to be implemented by subclasses"""
        raise NotImplementedError("Subclasses must implement build_page method")
    
    def show(self):
        """Show the page"""
        if not self.page_frame:
            self.create_page()
        
        if not self.is_visible:
            self.page_frame.pack(fill='both', expand=True)
            self.is_visible = True
            self.on_show()
    
    def hide(self):
        """Hide the page"""
        if self.is_visible and self.page_frame:
            self.page_frame.pack_forget()
            self.is_visible = False
            self.on_hide()
    
    def on_show(self):
        """Called when page is shown - can be overridden by subclasses"""
        pass
    
    def on_hide(self):
        """Called when page is hidden - can be overridden by subclasses"""
        pass
    
    def destroy_page(self):
        """Destroy the page frame"""
        if self.page_frame:
            self.page_frame.destroy()
            self.page_frame = None
            self.is_visible = False