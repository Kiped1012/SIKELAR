#!/usr/bin/env python3
"""
SIKELAR - Sistem Informasi Pengelompokan Anggaran dan Rekening
Main entry point for the application
"""

import tkinter as tk
from gui.app_gui import BOSBudgetAnalyzer

def main():
    """Main function to run the SIKELAR application"""
    root = tk.Tk()
    app = BOSBudgetAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()