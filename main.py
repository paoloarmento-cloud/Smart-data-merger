#!/usr/bin/env python3
"""
Excel/CSV Merger Tool - Main Entry Point
Portfolio project demonstrating data processing and GUI development
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox
import logging

# Add current directory to path for imports
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

try:
    from interface import main as gui_main
except ImportError as e:
    sys.exit(1)

def check_dependencies():
    """Check if all required packages are available"""
    required_packages = {
        'pandas': 'pandas',
        'openpyxl': 'openpyxl', 
        'fuzzywuzzy': 'fuzzywuzzy',
        'tkinter': 'tkinter (should be built-in)'
    }
    
    missing = []
    
    try:
        import pandas
    except ImportError:
        missing.append('pandas')
    
    try:
        import openpyxl
    except ImportError:
        missing.append('openpyxl')
        
    try:
        from fuzzywuzzy import fuzz
    except ImportError:
        missing.append('fuzzywuzzy')
    
    try:
        import tkinter
    except ImportError:
        missing.append('tkinter')
    
    if missing:
        error_msg = "Missing required packages:\n\n"
        for pkg in missing:
            error_msg += f"  • {pkg}\n"
        error_msg += "\nPlease install missing packages using:\n"
        error_msg += f"pip install {' '.join(missing)}"
        
        # Try to show GUI error if tkinter available, otherwise print
        try:
            root = tk.Tk()
            root.withdraw()  # Hide main window
            messagebox.showerror("Missing Dependencies", error_msg)
        except:
            print(error_msg)
        
        return False
    
    return True

def setup_logging():
    """Setup logging configuration"""
    log_dir = os.path.join(os.path.expanduser("~"), "ExcelMerger")
    os.makedirs(log_dir, exist_ok=True)
    
    log_file = os.path.join(log_dir, "merger.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("Excel/CSV Merger Tool starting...")
    logger.info(f"Log file: {log_file}")

def main():
    """Main application entry point"""
    
    # Setup logging
    setup_logging()
    logger = logging.getLogger(__name__)
    
    try:
        # Check dependencies
        if not check_dependencies():
            logger.error("Dependency check failed")
            return 1
        
        logger.info("Dependencies check passed")
        
        # Launch GUI
        logger.info("Launching GUI...")
        gui_main()
        
        logger.info("Application closed normally")
        return 0
        
    except KeyboardInterrupt:
        logger.info("Application interrupted by user")
        return 0
        
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        
        # Try to show error dialog
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Unexpected Error", 
                f"An unexpected error occurred:\n\n{str(e)}\n\nCheck the log file for details."
            )
        except:
            print(f"Error: {str(e)}")
        
        return 1

if __name__ == "__main__":
    sys.exit(main())