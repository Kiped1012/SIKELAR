"""
Main entry point untuk aplikasi SIKELAR
Pastikan file ini dinamakan main.py atau sikelar_main.py
"""

import tkinter as tk
import os
import sys
import ctypes
from gui.main_app import SikelarMainApp  # Import class utama

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        print(f"Running from PyInstaller bundle: {base_path}")
    except Exception:
        base_path = os.path.abspath(".")
        print(f"Running from source directory: {base_path}")
    
    full_path = os.path.join(base_path, relative_path)
    print(f"Looking for resource: {full_path}")
    print(f"Resource exists: {os.path.exists(full_path)}")
    
    # Debug: list available files in base directory
    try:
        available_files = os.listdir(base_path)
        print(f"Available files in {base_path}: {available_files[:10]}...")  # Show first 10 files
    except Exception as e:
        print(f"Could not list directory contents: {e}")
    
    return full_path

def set_window_icon(root):
    """Set window icon with multiple fallback options"""
    icon_names = ["sikelar_logo3.ico", "icon.ico", "app.ico"]
    
    for icon_name in icon_names:
        try:
            icon_path = resource_path(icon_name)
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
                print(f"✓ Window icon loaded successfully from: {icon_path}")
                return icon_path
            else:
                print(f"✗ Icon file not found: {icon_path}")
        except Exception as e:
            print(f"✗ Could not load icon {icon_name}: {e}")
    
    print("⚠ No icon could be loaded, using default")
    return None

def set_taskbar_icon():
    """Set taskbar icon for Windows"""
    try:
        # Set taskbar icon untuk Windows
        if sys.platform.startswith('win'):
            import ctypes
            # Mendapatkan handle aplikasi
            myappid = 'sikelar.app.1.0'  # arbitrary string
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            print("✓ Taskbar app ID set successfully")
    except Exception as e:
        print(f"⚠ Could not set taskbar icon: {e}")

def main():
    """Main function to run SIKELAR application"""
    # Set taskbar icon sebelum membuat window
    set_taskbar_icon()
    
    root = tk.Tk()
    
    # Set window icon
    icon_path = set_window_icon(root)
    
    # Jika icon berhasil dimuat, coba set sebagai default icon untuk semua window
    if icon_path and os.path.exists(icon_path):
        try:
            root.iconbitmap(default=icon_path)
            print("✓ Default icon set for all windows")
        except Exception as e:
            print(f"⚠ Could not set default icon: {e}")
    
    # Set window properties
    root.title("SIKELAR - Sistem Informasi Pengelompokan Anggaran")
    root.geometry("1200x800")
    root.minsize(1000, 700)
    root.configure(bg='#f8f9fa')
    
    # Tambahan untuk memastikan icon tetap ada
    root.wm_iconbitmap(bitmap=icon_path if icon_path and os.path.exists(icon_path) else "")
    
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