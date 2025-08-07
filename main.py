"""
Main entry point untuk aplikasi SIKELAR
Pastikan file ini dinamakan main.py atau sikelar_main.py
"""

import tkinter as tk
import os
import sys
from gui.main_app import SikelarMainApp  # Import class utama

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def main():
    """Main function to run SIKELAR application"""
    root = tk.Tk()
    
    # Set window icon (untuk taskbar saat aplikasi berjalan)
    try:
        icon_path = resource_path("sikelar_logo3.ico")
        root.iconbitmap(icon_path)
        print(f"Icon loaded successfully from: {icon_path}")
    except Exception as e:
        print(f"Could not load window icon: {e}")
    
    # Set window properties
    root.title("SIKELAR - Sistem Informasi Pengelompokan Anggaran")
    root.geometry("1200x800")
    root.minsize(1000, 700)
    root.configure(bg='#f8f9fa')
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Create and run application
    app = SikelarMainApp(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("Application terminated by user")
    except Exception as e:
        print(f"Application error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()