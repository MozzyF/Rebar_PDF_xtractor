import os
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import json
import csv
from collections import defaultdict
import pandas as pd
from datetime import datetime
import time

class DrawingInfo:
    """
    Represents information extracted from a PDF drawing.
    Stores drawing number, revision, title, and total weight.
    """
    def __init__(self, drawing_number="", revision="", title="", total_weight=0):
        self.drawing_number = drawing_number
        self.revision = revision
        self.title = title
        self.total_weight = total_weight

    def to_dict(self):
        """Convert drawing info to dictionary format for CSV/JSON export."""
        return {
            "Drawing Number": self.drawing_number,
            "Revision": self.revision,
            "Title": self.title,
            "Total Weight (kg)": self.total_weight
        }

class PDFAnalyzerGUI:
    """
    Main GUI application for analyzing PDF drawings and extracting weight information.
    Features a modern interface with progress tracking, real-time updates, and educational facts.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Rebar PDF to CSV")
        self.root.geometry("1200x800")  # Increased initial window size
        self.root.minsize(1000, 700)    # Increased minimum window size
        
        # Educational facts about rebar to display during processing
        self.rebar_facts = [
            "Rebar's ridged surface helps it bond better with concrete.",
            "The 'R' in rebar sizes (like R10) represents the bar's diameter in millimeters.",
            "Epoxy-coated rebar is used in corrosive environments to prevent rust.",
            "Rebar was first patented in 1849 by Joseph Monier.",
            "Modern rebar is made from recycled steel, making it environmentally friendly.",
            "The Romans used bronze and brass bars as ancient rebar in their concrete.",
            "Stainless steel rebar can last over 100 years without corrosion.",
            "Rebar increases concrete's tensile strength by up to 40,000 PSI.",
            "Glass fiber reinforced polymer (GFRP) rebar is becoming popular for its corrosion resistance.",
            "The world's annual rebar production exceeds 150 million tonnes."
        ]
        self.current_fact_index = 0
        self.fact_update_interval = 15000  # 15 seconds in milliseconds
        
        # Configure modern UI styles
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#2196F3")
        style.configure("TLabel", padding=6)
        style.configure("Header.TLabel", font=("Helvetica", 14, "bold"))
        style.configure("Fact.TLabel", font=("Helvetica", 10, "italic"))
        style.configure("Timer.TLabel", font=("Helvetica", 12, "bold"))
        style.configure("TLabelframe", padding=10)
        style.configure("TLabelframe.Label", font=("Helvetica", 10, "bold"))
        
        # Create main container with padding
        main_container = ttk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header with logo/icon (reduced vertical padding)
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        header = ttk.Label(header_frame, text="PDF Drawing Weight Analyzer", 
                          style="Header.TLabel")
        header.pack(side=tk.LEFT)
        
        # Directory selection frame (reduced vertical padding)
        dir_frame = ttk.LabelFrame(main_container, text="Directory Selection", padding=10)
        dir_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Input directory selection
        input_frame = ttk.Frame(dir_frame)
        input_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(input_frame, text="Input Directory:", width=15).pack(side=tk.LEFT)
        self.input_dir_var = tk.StringVar()
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_dir_var)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.select_input_dir, width=10).pack(side=tk.RIGHT)
        
        # Output directory selection
        output_frame = ttk.Frame(dir_frame)
        output_frame.pack(fill=tk.X)
        ttk.Label(output_frame, text="Output Directory:", width=15).pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar()
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="Browse", command=self.select_output_dir, width=10).pack(side=tk.RIGHT)
        
        # Progress frame with fixed height (reduced height and padding)
        progress_frame = ttk.LabelFrame(main_container, text="Progress", padding=5)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Fixed height container for progress bar (reduced height)
        progress_container = ttk.Frame(progress_frame, height=100)  # Increased height for facts
        progress_container.pack(fill=tk.X)
        progress_container.pack_propagate(False)
        
        # Progress bar with fixed height
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_container, variable=self.progress_var, 
                                          maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(5, 2))
        
        # Add facts label
        self.fact_var = tk.StringVar(value="Did you know? " + self.rebar_facts[0])
        self.fact_label = ttk.Label(progress_container, textvariable=self.fact_var,
                                  style="Fact.TLabel", wraplength=800)
        self.fact_label.pack(fill=tk.X, pady=(2, 2))
        
        # Add estimated time remaining
        self.time_var = tk.StringVar(value="Estimated time remaining: --:--")
        self.time_label = ttk.Label(progress_container, textvariable=self.time_var,
                                  style="Timer.TLabel")
        self.time_label.pack(fill=tk.X)
        
        # Status label with fixed height
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(progress_container, textvariable=self.status_var)
        self.status_label.pack(fill=tk.X)
        
        # Results frame (adjusted padding)
        results_frame = ttk.LabelFrame(main_container, text="Results", padding=5)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create Treeview with fixed columns
        self.tree = ttk.Treeview(results_frame, columns=("Drawing", "Revision", "Title", "Weight", "Pages"), 
                                show="headings", height=15)
        
        # Configure column widths
        self.tree.column("Drawing", width=100, minwidth=80, anchor=tk.CENTER)
        self.tree.column("Revision", width=80, minwidth=60, anchor=tk.CENTER)
        self.tree.column("Title", width=500, minwidth=200, anchor=tk.W)
        self.tree.column("Weight", width=100, minwidth=80, anchor=tk.CENTER)
        self.tree.column("Pages", width=100, minwidth=80, anchor=tk.CENTER)
        
        # Configure headings
        self.tree.heading("Drawing", text="Drawing #")
        self.tree.heading("Revision", text="Revision")
        self.tree.heading("Title", text="Title")
        self.tree.heading("Weight", text="Weight (KG)")
        self.tree.heading("Pages", text="Pages")
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Grid layout for tree and scrollbars
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 2))
        y_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        x_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Configure grid weights
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Button frame at bottom (ensure it's always visible)
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Open Results button (initially disabled)
        self.open_results_button = ttk.Button(button_frame, text="Open Results", 
                                            command=self.open_results, width=20,
                                            state='disabled')
        self.open_results_button.pack(side=tk.RIGHT, padx=5)
        
        # Start button with fixed width
        self.start_button = ttk.Button(button_frame, text="Start Analysis", 
                                     command=self.start_analysis, width=20)
        self.start_button.pack(side=tk.RIGHT, padx=5)
        
        # Store the latest Excel file path
        self.latest_excel_path = None
        
        # Bind window resize event
        self.root.bind('<Configure>', self.on_window_resize)
        
        # Initial layout update
        self.update_layout()

    def on_window_resize(self, event):
        """Handle window resize events"""
        if event.widget == self.root:
            self.update_layout()

    def update_layout(self):
        """Update layout to maintain proper sizing"""
        # Update tree column widths based on window size
        tree_width = self.tree.winfo_width()
        if tree_width > 0:
            self.tree.column("Title", width=tree_width - 300)  # Adjust for other columns

    def update_progress(self, value, status):
        """Update progress bar and status with smooth animation"""
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update_idletasks()

    def update_results(self, processed_drawings):
        """Update results tree with smooth animation"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add new items with alternating colors
        for drawing_number, drawings in processed_drawings.items():
            for drawing in drawings:
                page_count = len(drawing['page_weights'])
                item = self.tree.insert("", tk.END, values=(
                    drawing_number,
                    drawing.get('revision', ''),
                    drawing['title'],
                    f"{drawing['weight']:.1f}",
                    f"{page_count} pages"
                ))
                # Alternate row colors
                if self.tree.index(item) % 2 == 0:
                    self.tree.tag_configure('evenrow', background='#f0f0f0')
                    self.tree.item(item, tags=('evenrow',))
        
        self.root.update_idletasks()

    def open_results(self):
        """Open the Excel results file"""
        if self.latest_excel_path and os.path.exists(self.latest_excel_path):
            try:
                os.startfile(self.latest_excel_path)  # Windows
            except AttributeError:
                try:
                    import subprocess
                    subprocess.call(['open', self.latest_excel_path])  # macOS
                except:
                    try:
                        subprocess.call(['xdg-open', self.latest_excel_path])  # Linux
                    except:
                        messagebox.showerror("Error", "Could not open the results file automatically. "
                                           f"Please open it manually at: {self.latest_excel_path}")
        else:
            messagebox.showerror("Error", "No results file available. Please run the analysis first.")

    def start_analysis(self):
        """Start analysis with visual feedback"""
        input_dir = self.input_dir_var.get()
        output_dir = self.output_dir_var.get()
        
        if not input_dir or not output_dir:
            messagebox.showerror("Error", "Please select both input and output directories.")
            return
        
        # Disable controls and update visual state
        self.start_button.state(['disabled'])
        self.input_entry.state(['disabled'])
        self.output_entry.state(['disabled'])
        self.open_results_button.state(['disabled'])  # Disable during analysis
        self.progress_var.set(0)
        self.status_var.set("Starting analysis...")
        
        try:
            # Process PDFs
            processed_drawings = self.process_pdfs(input_dir, output_dir)
            
            # Update results display
            self.update_results(processed_drawings)
            
            # Show completion message
            messagebox.showinfo("Success", "Analysis completed successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.open_results_button.state(['disabled'])  # Disable on error
        
        finally:
            # Re-enable controls
            self.start_button.state(['!disabled'])
            self.input_entry.state(['!disabled'])
            self.output_entry.state(['!disabled'])
            self.status_var.set("Ready")

    def cycle_fact(self):
        """Cycle to the next rebar fact"""
        self.current_fact_index = (self.current_fact_index + 1) % len(self.rebar_facts)
        self.fact_var.set("Did you know? " + self.rebar_facts[self.current_fact_index])
        if hasattr(self, 'processing') and self.processing:
            self.root.after(self.fact_update_interval, self.cycle_fact)

    def update_time_remaining(self, start_time, total_files, processed_files):
        """Update the estimated time remaining"""
        try:
            if not total_files or processed_files <= 0:
                self.time_var.set("Estimated time remaining: Calculating...")
                if hasattr(self, 'processing') and self.processing:
                    self.root.after(1000, lambda: self.update_time_remaining(start_time, total_files, processed_files))
                return
            
            elapsed_time = max(0.1, time.time() - start_time)  # Ensure we don't divide by zero
            files_remaining = max(0, total_files - processed_files)
            
            # Protect against division by zero
            try:
                time_per_file = elapsed_time / processed_files
                estimated_remaining = files_remaining * time_per_file
            except (ZeroDivisionError, ValueError):
                self.time_var.set("Estimated time remaining: Calculating...")
                if hasattr(self, 'processing') and self.processing:
                    self.root.after(1000, lambda: self.update_time_remaining(start_time, total_files, processed_files))
                return
            
            minutes = int(estimated_remaining // 60)
            seconds = int(estimated_remaining % 60)
            self.time_var.set(f"Estimated time remaining: {minutes:02d}:{seconds:02d}")
            
            if hasattr(self, 'processing') and self.processing:
                self.root.after(1000, lambda: self.update_time_remaining(start_time, total_files, processed_files))
        except Exception as e:
            # If any error occurs, show calculating message and continue
            self.time_var.set("Estimated time remaining: Calculating...")
            if hasattr(self, 'processing') and self.processing:
                self.root.after(1000, lambda: self.update_time_remaining(start_time, total_files, processed_files))

    def process_pdfs(self, input_dir, output_dir):
        """Process all PDFs in the input directory"""
        try:
            # Create output directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)
            
            # Get all PDF files
            pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                raise ValueError("No PDF files found in the input directory.")
            
            # Track processed drawings and duplicates
            processed_drawings = {}
            duplicate_groups = []
            total_files = len(pdf_files)
            start_time = time.time()
            self.processing = True
            
            # Start cycling facts and updating time
            self.cycle_fact()
            self.update_time_remaining(start_time, total_files, 0)
            
            # Process each PDF
            for index, pdf_file in enumerate(pdf_files, 1):
                pdf_path = os.path.join(input_dir, pdf_file)
                self.update_progress((index / total_files) * 100, f"Processing: {pdf_file}")
                self.update_time_remaining(start_time, total_files, index)
                
                # Extract drawing info
                drawing_number, title = extract_drawing_info(pdf_file)
                if drawing_number and title:
                    # Extract revision from the PDF text
                    full_text, pages_text = extract_text_from_pdf(pdf_path)
                    if full_text:  # Check if text extraction was successful
                        revision = extract_revision(full_text)
                    else:
                        revision = ""
                    
                    # Process the PDF and get weights
                    weight, page_weights = extract_total_weight(pdf_path, output_dir)
                    
                    # Initialize list for this drawing number if not exists
                    if drawing_number not in processed_drawings:
                        processed_drawings[drawing_number] = []
                    
                    # Add the drawing info
                    processed_drawings[drawing_number].append({
                        'title': title,
                        'revision': revision,
                        'weight': weight,
                        'page_weights': page_weights,
                        'filename': pdf_file
                    })
            
            # Check for duplicates (any drawings with the same number)
            for drawing_number, drawings in processed_drawings.items():
                if len(drawings) > 1:
                    # Sort drawings by weight for easier comparison
                    drawings.sort(key=lambda x: x['weight'])
                    duplicate_groups.append((drawing_number, drawings))
            
            # If duplicates found, show resolution dialog
            if duplicate_groups:
                dialog = DuplicateResolutionDialog(self.root, duplicate_groups, input_dir)
                selections = dialog.get_selections()
                
                # Update processed_drawings with user selections
                for drawing_number, selected_drawings in selections.items():
                    processed_drawings[drawing_number] = selected_drawings
            
            # Save results to Excel
            self.save_to_excel(processed_drawings, output_dir)
            
            self.processing = False
            return processed_drawings
            
        except Exception as e:
            self.processing = False
            raise Exception(f"Error processing PDFs: {str(e)}")

    def save_to_excel(self, processed_drawings, output_dir):
        """Save results to Excel file"""
        try:
            rows = []
            for drawing_number, drawings in processed_drawings.items():
                for drawing in drawings:
                    row = {
                        'Drawing Number': drawing_number,
                        'Revision': drawing.get('revision', ''),
                        'Title': drawing['title'],
                        'Total Weight (KG)': drawing['weight'],
                        'Page Weights': str(drawing['page_weights']),
                        'Filename': drawing['filename']
                    }
                    rows.append(row)
            
            df = pd.DataFrame(rows)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_path = os.path.join(output_dir, f'drawing_weights_{timestamp}.xlsx')
            df.to_excel(excel_path, index=False)
            
            # Store the path and enable the Open Results button
            self.latest_excel_path = excel_path
            self.open_results_button.state(['!disabled'])
            
        except Exception as e:
            raise Exception(f"Error saving to Excel: {str(e)}")

    def select_input_dir(self):
        """Open dialog to select input directory"""
        directory = filedialog.askdirectory(title="Select Input Directory")
        if directory:
            self.input_dir_var.set(directory)

    def select_output_dir(self):
        """Open dialog to select output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)

class DuplicateResolutionDialog:
    def __init__(self, parent, duplicates, input_dir):
        """
        duplicates: list of tuples (drawing_number, list of drawing dictionaries)
        input_dir: path to the input directory containing PDFs
        """
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Resolve Duplicate Drawings")
        self.dialog.geometry("800x800")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.selected_drawings = {}  # Store user selections
        self.duplicates = self.sort_duplicates(duplicates)  # Sort duplicates by similarity
        self.input_dir = input_dir
        
        # Create main frame
        main_frame = ttk.Frame(self.dialog, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add summary header
        total_groups = len(duplicates)
        total_files = sum(len(drawings) for _, drawings in duplicates)
        summary_text = f"Found {total_groups} drawing number(s) with multiple versions ({total_files} total files)"
        ttk.Label(main_frame, 
                 text=summary_text,
                 style="Summary.TLabel").pack(pady=(0, 5))
        
        # Add instructions
        ttk.Label(main_frame, 
                 text="Select which drawings to include in the final output (you can select multiple):", 
                 wraplength=700).pack(pady=(0, 5))
        
        # Create canvas and scrollbar for scrolling
        self.canvas = tk.Canvas(main_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Configure canvas scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            self._on_frame_configure
        )
        
        # Create window in canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind canvas resize
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        
        # Bind mouse wheel events
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Configure styles
        self.configure_styles()
        
        # Add duplicate groups
        self.create_duplicate_groups()
        
        # Add buttons
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(button_frame, text="Confirm Selections", 
                  command=self.confirm_selections).pack(side="right", padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side="right", padx=5)

    def sort_duplicates(self, duplicates):
        """Sort duplicates by number of versions and then by weight similarity"""
        sorted_duplicates = []
        for drawing_number, drawings in duplicates:
            # Sort drawings by weight
            drawings = sorted(drawings, key=lambda x: x['weight'])
            
            # Calculate weight differences between consecutive drawings
            max_diff = 0
            for i in range(len(drawings)-1):
                diff = abs(drawings[i]['weight'] - drawings[i+1]['weight'])
                max_diff = max(max_diff, diff)
            
            # Store tuple of (drawing_number, drawings, num_duplicates, max_weight_difference)
            sorted_duplicates.append((drawing_number, drawings, len(drawings), max_diff))
        
        # Sort first by number of duplicates (descending), then by weight difference (ascending)
        sorted_duplicates.sort(key=lambda x: (-x[2], x[3]))
        
        # Return just the drawing_number and drawings
        return [(num, drw) for num, drw, _, _ in sorted_duplicates]

    def configure_styles(self):
        """Configure ttk styles for the dialog"""
        style = ttk.Style()
        style.configure("Title.TLabel", wraplength=600, font=("Helvetica", 9))
        style.configure("Weight.TLabel", font=("Helvetica", 9, "bold"))
        style.configure("Summary.TLabel", font=("Helvetica", 10, "bold"))
        style.configure("Group.TLabelframe", padding=5)
        style.configure("Group.TLabelframe.Label", font=("Helvetica", 9, "bold"))

    def _on_frame_configure(self, event=None):
        """Reset the scroll region to encompass the inner frame"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """When canvas is resized, resize the inner frame to match"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if event.num == 5 or event.delta < 0:  # Scroll down
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:  # Scroll up
            self.canvas.yview_scroll(-1, "units")

    def create_duplicate_groups(self):
        """Create a group of checkboxes for each set of duplicates"""
        for drawing_number, drawings in self.duplicates:
            # Create frame for this group
            group_frame = ttk.LabelFrame(self.scrollable_frame, 
                                       text=f"Drawing Number: {drawing_number} ({len(drawings)} versions)", 
                                       padding="5",
                                       style="Group.TLabelframe")
            group_frame.pack(fill="x", padx=5, pady=2)
            
            # Calculate weight differences
            weight_diffs = []
            for i in range(len(drawings)-1):
                diff = abs(drawings[i]['weight'] - drawings[i+1]['weight'])
                weight_diffs.append(f"{diff:.2f}KG")
            
            # Add weight difference info if available
            if weight_diffs:
                diff_text = f"Weight differences: {' / '.join(weight_diffs)}"
                ttk.Label(group_frame, text=diff_text, style="Weight.TLabel").pack(fill="x", pady=(0, 5))
            
            # List to store checkboxes for this group
            self.selected_drawings[drawing_number] = []
            
            # Add checkboxes for each drawing
            for idx, drawing in enumerate(drawings):
                drawing_frame = ttk.Frame(group_frame)
                drawing_frame.pack(fill="x", pady=1)
                
                # Create checkbox variable
                var = tk.BooleanVar(value=True)
                self.selected_drawings[drawing_number].append((var, drawing))
                
                # Top frame for checkbox and basic info
                top_frame = ttk.Frame(drawing_frame)
                top_frame.pack(fill="x")
                
                # Left side with checkbox and revision
                left_frame = ttk.Frame(top_frame)
                left_frame.pack(side="left", fill="x", expand=True)
                
                # Create checkbox with revision
                checkbox = ttk.Checkbutton(
                    left_frame, 
                    text=f"Rev: {drawing.get('revision', 'N/A')}",
                    variable=var
                )
                checkbox.pack(side="left", padx=(0, 5))
                
                # Add total weight
                ttk.Label(left_frame, 
                         text=f"Weight: {drawing['weight']:.2f} KG",
                         style="Weight.TLabel").pack(side="left", padx=(0, 5))
                
                # Format page weights as compact text
                page_weights = drawing.get('page_weights', [])
                if page_weights:
                    page_text = " / ".join([f"P{p}:{w:.2f}" for p, w in page_weights])
                    ttk.Label(left_frame, text=page_text).pack(side="left", padx=(0, 5))
                
                # Right side with PDF button and filename
                right_frame = ttk.Frame(top_frame)
                right_frame.pack(side="right")
                
                # Add View PDF button
                ttk.Button(
                    right_frame, 
                    text="View PDF", 
                    command=lambda f=drawing['filename']: self.open_pdf(f)
                ).pack(side="right", padx=2)
                
                # Add filename (shortened if too long)
                filename = drawing['filename']
                if len(filename) > 40:
                    filename = filename[:37] + "..."
                ttk.Label(right_frame, text=filename).pack(side="right", padx=2)
                
                # Add title below
                ttk.Label(drawing_frame, 
                         text=drawing['title'], 
                         style="Title.TLabel").pack(fill="x", pady=(1, 0))
                
                if idx < len(drawings) - 1:  # Don't add separator after last item
                    ttk.Separator(group_frame, orient="horizontal").pack(fill="x", pady=1)

    def open_pdf(self, filename):
        """Open the PDF file in the default viewer"""
        try:
            pdf_path = os.path.join(self.input_dir, filename)
            if os.path.exists(pdf_path):
                os.startfile(pdf_path)  # Windows
            else:
                messagebox.showerror("Error", f"Could not find PDF file: {filename}")
        except AttributeError:
            try:
                import subprocess
                subprocess.call(['open', pdf_path])  # macOS
            except:
                try:
                    subprocess.call(['xdg-open', pdf_path])  # Linux
                except:
                    messagebox.showerror("Error", f"Could not open PDF: {filename}")

    def confirm_selections(self):
        """Store the selected drawings and close the dialog"""
        self.final_selections = {}
        for drawing_number, selections in self.selected_drawings.items():
            selected_drawings = [drawing for var, drawing in selections if var.get()]
            if selected_drawings:
                self.final_selections[drawing_number] = selected_drawings
        self.dialog.destroy()

    def get_selections(self):
        """Wait for dialog to close and return selections"""
        self.dialog.wait_window()
        return getattr(self, 'final_selections', {})

def select_directory(title):
    """Open a file dialog to select a directory."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    directory = filedialog.askdirectory(title=title)
    return directory

def extract_drawing_info(filename):
    """Extract drawing number and title from filename"""
    # Try different patterns for drawing number
    drawing_patterns = [
        r'-(\d{4})(?:_|$)',  # Matches -2000_ or -2000
        r'(\d{4})(?:_|$)',   # Matches 2000_ or 2000
        r'DR-S-(\d{4})',     # Matches DR-S-2000
        r'FN-DR-S-(\d{4})',  # Matches FN-DR-S-2000
        r'Drawing\s*No\.\s*(\d{4})'  # Matches Drawing No. 2000
    ]
    
    drawing_number = ""
    for pattern in drawing_patterns:
        match = re.search(pattern, filename)
        if match:
            drawing_number = match.group(1)
            break
    
    # Extract title from filename
    # Remove file extension and common prefixes
    title = filename.replace('.pdf', '')
    
    # Remove drawing number and revision patterns
    title = re.sub(r'-\d{4}(?:_|$)', '', title)
    title = re.sub(r'DR-S-\d{4}', '', title)
    title = re.sub(r'FN-DR-S-\d{4}', '', title)
    title = re.sub(r'Drawing\s*No\.\s*\d{4}', '', title)
    
    # Remove revision patterns
    title = re.sub(r'_C\d{2}(?:_|$)', '', title)
    title = re.sub(r'_Construction_C\d{2}(?:_|$)', '', title)
    title = re.sub(r'_BBS_Construction_C\d{2}(?:_|$)', '', title)
    
    # Remove common prefixes and clean up
    title = re.sub(r'^1055-ACE-[A-Z]{2}-[0-9]{2}-[A-Z]{2}-S-', '', title)
    title = re.sub(r'^1055-ACE-[A-Z]{2}-FN-DR-S-', '', title)
    title = re.sub(r'^BBS_', '', title)
    title = re.sub(r'^_', '', title)
    title = re.sub(r'_+', ' ', title)
    title = title.strip()
    
    return drawing_number, title

def extract_revision(text):
    """Extract revision number from text."""
    patterns = [
        r'[_\s]C\d{2}\b',
        r'\bC\d{2}\b',
        r'REV[.:]\s*([A-Z0-9]+)',
        r'REVISION[.:]\s*([A-Z0-9]+)',
        r'Rev\.\s*([A-Z0-9]+)',
        r'Revision\s*([A-Z0-9]+)'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            if pattern.startswith(r'[_\s]C') or pattern.startswith(r'\bC'):
                return match.group(0).strip()
            return match.group(1)
    return ""

def extract_title(text, filename):
    """Extract title from text or filename."""
    # Try to extract from filename first
    filename_parts = filename.split('_')
    if len(filename_parts) > 1:
        # Get the part that looks like a title (usually the second part)
        potential_title = filename_parts[1].replace('-', ' ').strip()
        if len(potential_title) > 10:  # Arbitrary minimum length for a title
            return potential_title
    
    # Look for common title patterns in the text
    patterns = [
        r'TITLE[.:]\s*(.+?)(?=\n|$)',
        r'DRAWING TITLE[.:]\s*(.+?)(?=\n|$)',
        r'(?<=\n)(?!REV|DATE|SCALE)([A-Z][A-Z\s]+(?:\s+\d+(?:\s*TO\s*\d+)?)+)(?=\n)'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    
    return ""

def extract_text_from_pdf(pdf_path):
    """Extract text and information from a single PDF file."""
    try:
        reader = PdfReader(pdf_path)
        # Store text from each page separately
        pages_text = []
        for page in reader.pages:
            pages_text.append(page.extract_text())
        # Also return combined text for other extractions
        full_text = "\n".join(pages_text)
        return full_text, pages_text
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return None, None

def extract_total_weight(pdf_path, output_dir):
    """
    Extract total weight from PDF using multiple strategies and patterns.
    Handles various formats including split numbers, line breaks, and different separators.
    
    Args:
        pdf_path: Path to the PDF file
        output_dir: Directory for debug output
        
    Returns:
        tuple: (total_weight, page_weights) where page_weights is a list of (page_num, weight) tuples
    """
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            total_weight = 0
            page_weights = []
            
            # Comprehensive set of regex patterns for weight extraction
            # Each pattern handles a different format or edge case
            weight_patterns = [
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
            
            # Process each page of the PDF
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                
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
                                    total_weight += weight
                                    page_weights.append((page_num + 1, weight))
                                    print(f"Found split number on page {page_num + 1}:")
                                    print(f"Whole part: {whole_part}")
                                    print(f"Decimal part: {decimal_part}")
                                    print(f"Extracted weight: {weight} KG")
                                    weight_found = True
                                    break
                    
                    # Strategy 2: Try all regex patterns
                    if not weight_found:
                        for pattern in weight_patterns:
                            matches = re.finditer(pattern, line, re.IGNORECASE | re.MULTILINE)
                            for match in matches:
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
                                                        print(f"Found additional digit on next line: {digit_match.group(1)}")
                                                weight = float(weight_str)
                                    
                                    total_weight += weight
                                    page_weights.append((page_num + 1, weight))
                                    print(f"Found weight on page {page_num + 1}: {weight} KG (using pattern: {pattern})")
                                    weight_found = True
                                    break
                                except ValueError:
                                    print(f"Warning: Could not convert weight value on page {page_num + 1}")
                                    continue
                            
                            if weight_found:
                                break
                    
                    if weight_found:
                        break
                
                if not weight_found:
                    print(f"No weight found on page {page_num + 1}")
                    # Print last 200 characters of text for debugging
                    print(f"Last 200 chars: {text[-200:]}")
            
            # Check if this PDF should be included in final output
            if total_weight > 0:
                print(f"Found weight for {os.path.basename(pdf_path)}: {total_weight} KG")
                print(f"Page weights: {page_weights}")
            else:
                print(f"EXCLUDED: {os.path.basename(pdf_path)} - No weight found")
            
            return total_weight, page_weights
            
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return 0, []

def save_detailed_weights(weights_data, output_dir):
    """
    Save detailed weight information to a CSV file for analysis.
    Includes filename, drawing number, weight, pattern used, and page number.
    """
    csv_path = os.path.join(output_dir, 'detailed_weights.csv')
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Filename', 'Drawing Number', 'Weight (kg)', 'Pattern Used', 'Page Number'])
        for item in weights_data:
            writer.writerow([
                item['filename'],
                item['drawing_number'],
                item['weight'],
                item['pattern'],
                item.get('page', 'N/A')
            ])

def process_directory(input_dir, output_dir):
    """
    Process all PDF files in the input directory and extract weight information.
    Creates debug logs and saves results in multiple formats.
    
    Args:
        input_dir: Directory containing PDF files
        output_dir: Directory for saving results and debug logs
        
    Returns:
        tuple: (processed_count, failed_count)
    """
    os.makedirs(output_dir, exist_ok=True)
    
    # Dictionary to store drawing information
    drawings = defaultdict(list)  # drawing number -> list of DrawingInfo objects
    processed = 0
    failed = 0
    all_weights_data = []

    # Create debug files for troubleshooting
    debug_path = os.path.join(output_dir, 'debug_output.txt')
    weight_debug_path = os.path.join(output_dir, 'weight_debug.txt')
    
    # Clear weight debug file
    with open(weight_debug_path, 'w', encoding='utf-8') as f:
        f.write("Weight Extraction Debug Log\n")
        f.write("=" * 50 + "\n")

    # Process each PDF file
    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(input_dir, filename)
            print(f"\nProcessing: {filename}")
            
            try:
                # Extract text from PDF
                full_text, pages_text = extract_text_from_pdf(pdf_path)
                if not full_text:
                    print(f"Failed to extract text from {filename}")
                    failed += 1
                    continue
                
                # Extract drawing information
                drawing_number, title = extract_drawing_info(filename)
                revision = extract_revision(full_text)
                
                # Extract weights
                total_weight, weights_data = extract_total_weight(pdf_path, output_dir)
                all_weights_data.extend(weights_data)
                
                # Create DrawingInfo object
                drawing_info = DrawingInfo(
                    drawing_number=drawing_number,
                    revision=revision,
                    title=title,
                    total_weight=total_weight
                )
                
                # Add to drawings dictionary
                if drawing_number:
                    drawings[drawing_number].append(drawing_info)
                
                processed += 1
                print(f"Successfully processed {filename}")
                
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
                failed += 1
    
    # Save results in multiple formats
    save_results(drawings, output_dir)
    save_detailed_weights(all_weights_data, output_dir)
    
    return processed, failed

def save_results(drawings, output_dir):
    """
    Save the extracted information as CSV and JSON files.
    CSV format for easy viewing, JSON format for programmatic processing.
    """
    # Convert to list of dictionaries
    results = [info.to_dict() for info in drawings.values()]
    
    # Save as CSV
    csv_path = os.path.join(output_dir, 'drawing_information.csv')
    if results:
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=results[0].keys())
            writer.writeheader()
            writer.writerows(results)
        
        # Also save as JSON for easier processing
        json_path = os.path.join(output_dir, 'drawing_information.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2)

def main():
    root = tk.Tk()
    app = PDFAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 