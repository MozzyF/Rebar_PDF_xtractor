# Rebar PDF to CSV

A Python application that analyzes rebar PDF drawings to extract weight information and exports it to CSV format. It provides a modern graphical interface with progress tracking, real-time updates, and educational facts about rebar.

## Features

- **Modern GUI Interface**: User-friendly interface with progress tracking and real-time updates
- **Educational Content**: Displays interesting facts about rebar during processing
- **Duplicate Detection**: Identifies and helps resolve duplicate drawings
- **Weight Extraction**: Extracts weight information from various PDF formats
- **Multiple Output Formats**: Saves results in CSV and Excel formats
- **Debug Tool**: Includes a separate debugger for troubleshooting weight extraction
- **PDF Preview**: Allows viewing PDFs directly from the duplicate resolution dialog

## Requirements

- Python 3.6 or higher
- PyPDF2 library
- tkinter (usually comes with Python)
- pandas (for Excel output)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/Rebar_PDF_to_CSV.git
cd Rebar_PDF_to_CSV
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the main application:
```bash
python pdf_text_extractor.py
```

2. For debugging weight extraction:
```bash
python debug_weight_extraction.py
```

## Main Application Features

- **Directory Selection**: Choose input directory with PDFs and output directory for results
- **Progress Tracking**: Real-time progress bar with estimated time remaining
- **Educational Facts**: Learn about rebar while processing
- **Results Display**: View extracted information in a sortable table
- **Duplicate Resolution**: Interactive dialog for handling duplicate drawings
- **PDF Preview**: View PDFs directly from the duplicate resolution dialog
- **Multiple Export Formats**: Results saved in CSV, JSON, and Excel formats

## Debug Tool Features

- **Pattern Analysis**: Shows which regex patterns matched in the PDF
- **Context Display**: Shows surrounding text for better understanding
- **Weight Extraction Details**: Displays detailed information about weight extraction
- **PDF Preview**: Open PDFs in browser for visual inspection
- **Copy Debug Output**: Easy copying of debug information

## Output Files

The application generates several output files:
- `drawing_weights_[timestamp].xlsx`: Main results in Excel format
- `drawing_information.csv`: Basic drawing information
- `drawing_information.json`: Detailed drawing information
- `detailed_weights.csv`: Detailed weight extraction data

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 