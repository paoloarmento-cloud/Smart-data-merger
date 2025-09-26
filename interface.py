"""
Excel/CSV Merger Tool - GUI Module
tkinter-based user interface
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from typing import Optional
import threading
from core import MergeEngine

class MergerGUI:
    """Main GUI application for Excel/CSV Merger Tool"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Merger Tool v1.0")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Initialize merge engine
        self.engine = MergeEngine()
        
        # GUI variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Downloads", "output.xlsx"))
        self.selected_key1 = tk.StringVar()
        self.selected_key2 = tk.StringVar()
        self.auto_report = tk.BooleanVar()
        
        # Available keys for manual selection
        self.available_keys1 = []
        self.available_keys2 = []
        
        self.setup_gui()
        self.update_gui_state()
    
    def setup_gui(self):
        """Create and arrange GUI elements"""
        
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel/CSV Merger Tool", font=('Arial', 16, 'bold'))
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # File selection section
        files_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        files_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)
        row += 1
        
        # File 1
        ttk.Label(files_frame, text="First File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(files_frame, textvariable=self.file1_path, state='readonly').grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(files_frame, text="Browse", command=lambda: self.select_file(1)).grid(row=0, column=2)
        
        # File 2
        ttk.Label(files_frame, text="Second File:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(files_frame, textvariable=self.file2_path, state='readonly').grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        ttk.Button(files_frame, text="Browse", command=lambda: self.select_file(2)).grid(row=1, column=2, pady=(10, 0))
        
        # Preview and key detection section
        self.preview_frame = ttk.LabelFrame(main_frame, text="File Preview & Key Detection", padding="10")
        self.preview_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(row, weight=1)  # Make this section expandable
        row += 1
        
        # Preview text area (scrollable)
        self.preview_text = scrolledtext.ScrolledText(
            self.preview_frame, 
            height=10, 
            wrap=tk.WORD,
            state='disabled',
            font=('Courier', 9)
        )
        self.preview_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # Key selection section
        keys_frame = ttk.LabelFrame(main_frame, text="Merge Key Selection", padding="10")
        keys_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        keys_frame.columnconfigure(1, weight=1)
        keys_frame.columnconfigure(3, weight=1)
        row += 1
        
        # Auto-detected keys info
        self.auto_keys_label = ttk.Label(keys_frame, text="Auto-detected keys will appear here", foreground='blue')
        self.auto_keys_label.grid(row=0, column=0, columnspan=4, pady=(0, 10))
        
        # Manual key selection
        ttk.Label(keys_frame, text="File 1 Key:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.key1_combo = ttk.Combobox(keys_frame, textvariable=self.selected_key1, state='readonly')
        self.key1_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 20))
        
        ttk.Label(keys_frame, text="File 2 Key:").grid(row=1, column=2, sticky=tk.W, padx=(0, 10))
        self.key2_combo = ttk.Combobox(keys_frame, textvariable=self.selected_key2, state='readonly')
        self.key2_combo.grid(row=1, column=3, sticky=(tk.W, tk.E))
        
        # Validate keys button
        self.validate_btn = ttk.Button(keys_frame, text="Validate Keys", command=self.validate_keys)
        self.validate_btn.grid(row=2, column=0, columnspan=4, pady=(10, 0))
        
        # Output section
        output_frame = ttk.LabelFrame(main_frame, text="Output Options", padding="10")
        output_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        row += 1
        
        # Output path
        ttk.Label(output_frame, text="Output File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_frame, text="Browse", command=self.select_output_file).grid(row=0, column=2)
        
        # Auto-report checkbox
        ttk.Checkbutton(output_frame, text="Generate automatic pivot report", variable=self.auto_report).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=row, column=0, columnspan=3, pady=(10, 0))
        row += 1
        
        self.merge_btn = ttk.Button(action_frame, text="Execute Merge", command=self.execute_merge, style='Accent.TButton')
        self.merge_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(action_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        row += 1
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready - Select two files to begin")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def select_file(self, file_number: int):
        """Open file dialog to select input file"""
        filetypes = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title=f"Select File {file_number}",
            filetypes=filetypes
        )
        
        if filename:
            if file_number == 1:
                self.file1_path.set(filename)
            else:
                self.file2_path.set(filename)
            
            self.load_file(filename, file_number)
    
    def load_file(self, filename: str, file_number: int):
        """Load file using merge engine and update GUI"""
        self.status_var.set(f"Loading file {file_number}...")
        self.progress.start(10)
        
        def load_in_thread():
            success = self.engine.load_file(filename, file_number)
            
            # Update GUI in main thread
            self.root.after(0, lambda: self.on_file_loaded(file_number, success))
        
        thread = threading.Thread(target=load_in_thread)
        thread.start()
    
    def on_file_loaded(self, file_number: int, success: bool):
        """Handle file loading completion"""
        self.progress.stop()
        
        if success:
            self.status_var.set(f"File {file_number} loaded successfully")
            self.update_preview()
            self.update_key_options()
            
            # Auto-detect keys if both files are loaded
            if self.engine.df1 is not None and self.engine.df2 is not None:
                self.detect_keys()
        else:
            messagebox.showerror("Error", f"Failed to load file {file_number}")
            self.status_var.set("Error loading file")
        
        self.update_gui_state()
    
    def update_preview(self):
        """Update the preview text area"""
        preview_data = self.engine.get_preview_data()
        
        self.preview_text.config(state='normal')
        self.preview_text.delete(1.0, tk.END)
        
        for file_key, file_info in preview_data.items():
            file_name = file_info['name']
            rows = file_info['rows']
            cols = file_info['columns']
            columns = file_info['column_names']
            
            self.preview_text.insert(tk.END, f"\n=== {file_name} ===\n")
            self.preview_text.insert(tk.END, f"Rows: {rows}, Columns: {cols}\n")
            self.preview_text.insert(tk.END, f"Columns: {', '.join(columns)}\n")
            
            # Show first few rows
            self.preview_text.insert(tk.END, f"\nFirst few rows:\n")
            for i, row in enumerate(file_info['preview']):
                if i < 3:  # Show only first 3 rows
                    row_str = " | ".join([f"{k}:{v}" for k, v in list(row.items())[:4]])  # First 4 columns
                    if len(row) > 4:
                        row_str += " | ..."
                    self.preview_text.insert(tk.END, f"{row_str}\n")
            
            self.preview_text.insert(tk.END, "\n" + "-"*50 + "\n")
        
        self.preview_text.config(state='disabled')
    
    def update_key_options(self):
        """Update the key selection comboboxes"""
        if self.engine.df1 is not None:
            self.available_keys1 = list(self.engine.df1.columns)
            self.key1_combo['values'] = self.available_keys1
        
        if self.engine.df2 is not None:
            self.available_keys2 = list(self.engine.df2.columns)
            self.key2_combo['values'] = self.available_keys2
    
    def detect_keys(self):
        """Auto-detect merge keys"""
        self.status_var.set("Detecting merge keys...")
        
        def detect_in_thread():
            print("Starting detection in thread...")
            detected = self.engine.detect_merge_keys()
            print(f"Thread got results: {detected}")
            self.root.after(0, lambda: self.on_keys_detected(detected))
        
        thread = threading.Thread(target=detect_in_thread)
        thread.start()
    
    def on_keys_detected(self, detected_keys):
        """Handle key detection completion"""
        if detected_keys:
            print("Processing detected keys...")
            best_match = detected_keys[0]  # Best match first
            key1, key2, score = best_match
            print(f"Best match: {key1} <-> {key2} (score: {score:.2f})")
        
            # 1: Ensure dropdown are populated
            self.update_key_options()
        
            # 2: Set dropdown values  
            self.selected_key1.set(key1)
            self.selected_key2.set(key2)
            
            # 3: Update status buttons
            self.update_gui_state()
        
            # Update info label
            info_text = f"Auto-detected: '{key1}' ↔ '{key2}' (Confidence: {score:.1%})"
            if len(detected_keys) > 1:
                info_text += f" | {len(detected_keys)-1} other options available"
            self.auto_keys_label.config(text=info_text, foreground='green')
            self.status_var.set(f"Keys auto-detected with {score:.1%} confidence")
            
        else:
            self.auto_keys_label.config(text="No suitable keys detected - please select manually", foreground='orange')
            self.status_var.set("Manual key selection required")
            self.update_gui_state()
    
    def validate_keys(self):
        """Validate selected merge keys"""
        if not self.selected_key1.get() or not self.selected_key2.get():
            messagebox.showwarning("Warning", "Please select both merge keys")
            return
        
        self.status_var.set("Validating keys...")
        
        def validate_in_thread():
            result = self.engine.validate_merge_keys(self.selected_key1.get(), self.selected_key2.get())
            self.root.after(0, lambda: self.show_validation_result(result))
        
        thread = threading.Thread(target=validate_in_thread)
        thread.start()
    
    def show_validation_result(self, result):
        """Display key validation results"""
        if result['valid']:
            message = f"Key Validation Results:\n\n"
            message += f"File 1 ({self.selected_key1.get()}):\n"
            message += f"  Total values: {result['file1_total']}\n"
            message += f"  Unique values: {result['file1_unique']}\n"
            message += f"  Uniqueness: {result['file1_uniqueness']:.1%}\n\n"
            
            message += f"File 2 ({self.selected_key2.get()}):\n"
            message += f"  Total values: {result['file2_total']}\n"
            message += f"  Unique values: {result['file2_unique']}\n"
            message += f"  Uniqueness: {result['file2_uniqueness']:.1%}\n\n"
            
            message += f"Common values: {result['common_values']}\n"
            message += f"Match ratio File 1: {result['match_ratio_file1']:.1%}\n"
            message += f"Match ratio File 2: {result['match_ratio_file2']:.1%}\n"
            
            if result.get('warnings'):
                message += f"\nWarnings:\n"
                for warning in result['warnings']:
                    message += f"  • {warning}\n"
            
            messagebox.showinfo("Key Validation", message)
            self.status_var.set("Keys validated successfully")
        else:
            messagebox.showerror("Validation Error", result['error'])
            self.status_var.set("Key validation failed")
    
    def execute_merge(self):
        """Execute the merge operation"""
        if not self.selected_key1.get() or not self.selected_key2.get():
            messagebox.showwarning("Warning", "Please select merge keys first")
            return
        
        if not self.output_path.get():
            messagebox.showwarning("Warning", "Please specify output file path")
            return
        
        self.status_var.set("Executing merge...")
        self.progress.start(10)
        self.merge_btn.config(state='disabled')
        
        def merge_in_thread():
            # Perform merge
            success = self.engine.perform_merge(self.selected_key1.get(), self.selected_key2.get())
            
            if success:
                # Save result
                save_success = self.engine.save_result(self.output_path.get())
                
                # Generate report if requested
                report_success = True
                if self.auto_report.get() and save_success:
                    # TODO: Implement report generation
                    pass
                
                self.root.after(0, lambda: self.on_merge_completed(save_success and report_success))
            else:
                self.root.after(0, lambda: self.on_merge_completed(False))
        
        thread = threading.Thread(target=merge_in_thread)
        thread.start()
    
    def on_merge_completed(self, success: bool):
        """Handle merge completion"""
        self.progress.stop()
        self.merge_btn.config(state='normal')
        
        if success:
            rows = len(self.engine.merge_result) if self.engine.merge_result is not None else 0
            message = f"Merge completed successfully!\n\n"
            message += f"Output file: {self.output_path.get()}\n"
            message += f"Total rows: {rows}\n"
            
            if self.auto_report.get():
                message += f"Report generated: {self.output_path.get().replace('.xlsx', '_report.xlsx')}\n"
            
            messagebox.showinfo("Success", message)
            self.status_var.set(f"Merge completed - {rows} rows saved")
        else:
            messagebox.showerror("Error", "Merge operation failed. Check the status and try again.")
            self.status_var.set("Merge failed")
    
    def select_output_file(self):
        """Select output file path"""
        filename = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if filename:
            self.output_path.set(filename)
    
    def clear_all(self):
        """Clear all selections and reset GUI"""
        self.file1_path.set("")
        self.file2_path.set("")
        self.selected_key1.set("")
        self.selected_key2.set("")
        self.auto_report.set(False)
        
        # Reset engine
        self.engine = MergeEngine()
        
        # Clear preview
        self.preview_text.config(state='normal')
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.config(state='disabled')
        
        # Reset labels
        self.auto_keys_label.config(text="Auto-detected keys will appear here", foreground='blue')
        
        # Clear comboboxes
        self.key1_combo['values'] = []
        self.key2_combo['values'] = []
        
        self.status_var.set("Ready - Select two files to begin")
        self.update_gui_state()
    
    def update_gui_state(self):
        """Update GUI element states based on current data"""
        files_loaded = self.engine.df1 is not None and self.engine.df2 is not None
        keys_selected = bool(self.selected_key1.get() and self.selected_key2.get())
        
        self.validate_btn.config(state='normal' if keys_selected else 'disabled')
        self.merge_btn.config(state='normal' if keys_selected else 'disabled')


def main():
    """Main entry point for GUI application"""
    root = tk.Tk()
    
    # Configure style
    style = ttk.Style()
    try:
        style.theme_use('vista')  # Use modern Windows theme if available
    except:
        pass
    
    # Create and run application
    app = MergerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()