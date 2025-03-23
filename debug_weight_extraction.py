import os
from PyPDF2 import PdfReader
import re
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
import json
import webbrowser

class WeightDebuggerGUI:
    """
    GUI tool for debugging PDF weight extraction logic.
    Allows testing individual files and viewing detailed extraction information.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Rebar PDF to CSV - Debug Tool")
        self.root.geometry("1200x800")
        
        # Configure style for consistent UI appearance
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=6)
        
        # Create main container with padding for better visual spacing
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # File selection frame with entry and buttons
        file_frame = ttk.LabelFrame(main_container, text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.select_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Open in Browser", command=self.open_in_browser).pack(side=tk.LEFT, padx=5)
        
        # Debug output frame with scrollable text widget
        debug_frame = ttk.LabelFrame(main_container, text="Debug Output", padding=10)
        debug_frame.pack(fill=tk.BOTH, expand=True)
        
        self.debug_text = scrolledtext.ScrolledText(debug_frame, wrap=tk.WORD, height=30)
        self.debug_text.pack(fill=tk.BOTH, expand=True)
        
        # Button frame for analysis and copy actions
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Analyze PDF", command=self.analyze_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Copy Debug Output", command=self.copy_debug_output).pack(side=tk.LEFT, padx=5)
        
        # Comprehensive set of regex patterns for weight extraction
        # Each pattern handles a different format or edge case
        self.weight_patterns = [
            # Standard formats with explicit "Total Weight" or "KG" units
            r'Total\s+Weight\s*=\s*(\d+[.,]?\d*)\s*KG',
            r'All\s+Bars\s+in\s+this\s+sheet\s+Total\s*(\d+[.,]?\d*)',
            r'Total\s*(\d+[.,]?\d*)',
            r'Total(\d+)[.,](\d+)',
            r'Total(\d+)',
            
            # Variations with different separators and optional units
            r'Total\s*(\d+)\s*[.,]\s*(\d+)',
            r'Total\s*(\d+)\s*[.,]?\s*(\d+)',
            r'Total\s*(\d+)\s*[.,]?\s*(\d+)\s*KG',
            r'Total\s*(\d+(?:\s+\d+)*[.,]?\d*)',
            r'Total\s*(\d+(?:\s+\d+)*[.,]?\d*)\s*KG',
            
            # Patterns for handling line breaks with different line endings
            r'Total\s*(\d+)\s*[.,]\s*\n\s*(\d+)',
            r'Total\s*(\d+)\s*[.,]\s*\r\s*(\d+)',
            r'Total\s*(\d+)\s*[.,]\s*\r\n\s*(\d+)',
            r'Total\s*(\d+)\s*[.,]\s*$[\n\r]*\s*(\d+)',
            
            # Generic patterns for line breaks and whitespace
            r'Total\s*(\d+)\s*[.,]\s*[\n\r]+\s*(\d+)',
            r'Total\s*(\d+)\s*[.,]\s*\s*(\d+)',
            
            # Patterns for numbers split across lines with periods
            r'Total\s*(\d+)\s*\.\s*\n\s*(\d+)',
            r'Total\s*(\d+)\s*\.\s*\r\s*(\d+)',
            r'Total\s*(\d+)\s*\.\s*\r\n\s*(\d+)',
            r'Total\s*(\d+)\s*\.\s*$[\n\r]*\s*(\d+)',
            
            # Special cases for numbers ending with periods
            r'Total\s*(\d+)\s*\.\s*$[\n\r]*\s*(\d+)',
            r'All\s+Bars\s+in\s+this\s+sheet\s+Total\s*(\d+)\s*\.\s*$[\n\r]*\s*(\d+)',
            
            # Specific patterns for the target PDF format
            r'Total\s*\n\s*(\d+)\s*\n\s*\.\s*(\d+)',
            r'Total\s*\n\s*(\d+)\s*\.\s*\n\s*(\d+)',
            r'Total\s*\n\s*(\d+)\s*\.\s*(\d+)',
            r'Total\s*\n\s*(\d+)\s*\.\s*(\d+)\s*Status',
            r'Total\s*\n\s*(\d+)\s*\.\s*(\d+)\s*Status\s+C\s*\(Resubmit\)'
        ]

    def select_file(self):
        """Opens a file dialog to select a PDF and automatically starts analysis."""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.analyze_pdf()  # Auto-analyze for better UX

    def open_in_browser(self):
        """Opens the selected PDF in the default web browser for visual inspection."""
        file_path = self.file_path_var.get()
        if file_path:
            webbrowser.open(file_path)

    def copy_debug_output(self):
        """Copies the debug output to clipboard and shows a confirmation message."""
        output = self.debug_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(output)
        messagebox.showinfo("Success", "Debug output copied to clipboard!")

    def analyze_pdf(self):
        """
        Analyzes the selected PDF file for weight information.
        Uses multiple patterns and strategies to extract weights, handling various formats
        and edge cases like split numbers and line breaks.
        """
        file_path = self.file_path_var.get()
        if not file_path:
            return
        
        self.debug_text.delete(1.0, tk.END)
        self.debug_text.insert(tk.END, f"=== PDF Analysis: {os.path.basename(file_path)} ===\n\n")
        
        try:
            with open(file_path, 'rb') as file:
                reader = PdfReader(file)
                
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    text = page.extract_text()
                    
                    # Show context for debugging
                    self.debug_text.insert(tk.END, f"=== Page {page_num + 1} ===\n")
                    self.debug_text.insert(tk.END, "Context (last 500 characters):\n")
                    self.debug_text.insert(tk.END, "-" * 50 + "\n")
                    self.debug_text.insert(tk.END, text[-500:] + "\n")
                    self.debug_text.insert(tk.END, "-" * 50 + "\n\n")
                    
                    # Process text line by line for better handling of line breaks
                    lines = text.split('\n')
                    
                    # Look for weight information using multiple strategies
                    weight_found = False
                    for i, line in enumerate(lines):
                        # Strategy 1: Look for "Total" on its own line
                        # This handles cases where the number is split across multiple lines
                        if line.strip().lower() == "total" and i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            # Look for a number that might end with a period
                            number_match = re.search(r'^(\d+)(?:\.)?$', next_line)
                            if number_match:
                                whole_part = number_match.group(1)
                                # Check next line for decimal part
                                if i + 2 < len(lines):
                                    decimal_line = lines[i + 2].strip()
                                    decimal_match = re.search(r'^\.?(\d+)', decimal_line)
                                    if decimal_match:
                                        decimal_part = decimal_match.group(1)
                                        weight = float(f"{whole_part}.{decimal_part}")
                                        self.debug_text.insert(tk.END, f"Found split number:\n")
                                        self.debug_text.insert(tk.END, f"Whole part: {whole_part}\n")
                                        self.debug_text.insert(tk.END, f"Decimal part: {decimal_part}\n")
                                        self.debug_text.insert(tk.END, f"Extracted weight: {weight} KG\n")
                                        self.debug_text.insert(tk.END, "=" * 50 + "\n\n")
                                        weight_found = True
                                        break
                        
                        # Strategy 2: Try all regex patterns
                        if not weight_found:
                            for pattern in self.weight_patterns:
                                matches = re.finditer(pattern, line, re.IGNORECASE | re.MULTILINE)
                                for match in matches:
                                    self.debug_text.insert(tk.END, f"Pattern: {pattern}\n")
                                    self.debug_text.insert(tk.END, f"Full match: {match.group(0)}\n")
                                    self.debug_text.insert(tk.END, f"Groups: {match.groups()}\n")
                                    
                                    try:
                                        # Handle different match group configurations
                                        if len(match.groups()) == 2:
                                            # Direct match of whole and decimal parts
                                            whole_part = match.group(1).replace(' ', '')
                                            decimal_part = match.group(2)
                                            weight = float(f"{whole_part}.{decimal_part}")
                                        else:
                                            # Handle single group matches with various formats
                                            weight_str = match.group(1).replace(' ', '')
                                            
                                            # Case 1: Line ends with a period
                                            if line.strip().endswith('.'):
                                                if i + 1 < len(lines):
                                                    next_line = lines[i + 1].strip()
                                                    decimal_match = re.search(r'^\s*(\d+)\s*$', next_line)
                                                    if decimal_match:
                                                        weight = float(f"{weight_str}.{decimal_match.group(1)}")
                                                    else:
                                                        weight = float(weight_str)
                                                else:
                                                    weight = float(weight_str)
                                            else:
                                                # Case 2: Look for decimal part in current line
                                                decimal_match = re.search(r'[.,]\s*(\d+)', line[match.end():])
                                                if decimal_match:
                                                    weight = float(f"{weight_str}.{decimal_match.group(1)}")
                                                else:
                                                    # Case 3: Check next line for additional digits
                                                    if i + 1 < len(lines):
                                                        next_line = lines[i + 1].strip()
                                                        digit_match = re.search(r'^\s*(\d)\s*$', next_line)
                                                        if digit_match:
                                                            weight_str += digit_match.group(1)
                                                            self.debug_text.insert(tk.END, f"Found additional digit on next line: {digit_match.group(1)}\n")
                                                    weight = float(weight_str)
                                        
                                        self.debug_text.insert(tk.END, f"Extracted weight: {weight} KG\n")
                                        self.debug_text.insert(tk.END, "=" * 50 + "\n\n")
                                        weight_found = True
                                        break
                                    except ValueError as e:
                                        self.debug_text.insert(tk.END, f"Error converting to float: {str(e)}\n")
                                        self.debug_text.insert(tk.END, "=" * 50 + "\n\n")
                                
                                if weight_found:
                                    break
                        
                        if weight_found:
                            break
                    
                    if not weight_found:
                        self.debug_text.insert(tk.END, "No weight found on this page\n")
                        self.debug_text.insert(tk.END, "=" * 50 + "\n\n")
                
                self.debug_text.insert(tk.END, "Analysis complete!\n")
                
        except Exception as e:
            self.debug_text.insert(tk.END, f"Error analyzing PDF: {str(e)}\n")

def main():
    """Entry point for the debugger application."""
    root = tk.Tk()
    app = WeightDebuggerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 