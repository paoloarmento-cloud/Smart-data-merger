# Smart Data Merger

A smart Excel/CSV merger with automatic key detection and format normalization. Solves data inconsistency issues that break standard merge tools.

## Problem Statement

When merging Excel or CSV files, common challenges include:
- **Format inconsistencies**: One file has `64356145.0` (float), another has `64356145` (integer)
- **Manual key selection**: Identifying which columns to join on requires domain knowledge
- **Data cleaning**: Whitespace, case sensitivity, and null values break matches
- **No-code requirement**: Non-technical users need GUI tools, not Python scripts

Standard tools (Excel VLOOKUP, basic pandas) fail when data formats don't match perfectly.

## Solution

Smart Data Merger automatically:
- **Detects merge keys** by analyzing column names and value overlap
- **Normalizes data** to handle format differences (floats vs integers, whitespace, etc.)
- **Validates matches** before merging with detailed statistics
- **Provides GUI** for users without programming experience

## Features

-  Automatic merge key detection with confidence scoring
-  Intelligent data normalization (handles numeric format differences)
-  Support for Excel (.xlsx, .xls) and CSV files
-  Real-time preview of loaded files
-  Merge validation with overlap statistics
-  Multiple merge types (left, right, inner, outer)
-  User-friendly GUI built with tkinter
-  Detailed logging for troubleshooting

## Installation

### Option 1: Download Executable (Windows)
Download the latest `.exe` file from the [Releases](../../releases) page. No Python installation required.

### Option 2: Run from Source

**Requirements:**
- Python 3.8 or higher

**Install dependencies:**
```bash
pip install -r requirements.txt
```

**Run the application:**
```bash
python main.py
```

## Usage

1. **Load Files**: Click "Browse" to select two Excel or CSV files
2. **Auto-Detection**: The tool automatically identifies potential merge keys
3. **Validate**: Review the suggested keys and validate match statistics
4. **Configure Output**: Choose output file location
5. **Merge**: Click "Execute Merge" to create the merged file


## Project Structure

```
smart-data-merger/
├── main.py              # Entry point and dependency check
├── core.py              # Merge engine and data processing logic
├── interface.py         # GUI implementation
├── requirements.txt     # Python dependencies
├── README.md
└── LICENSE
```

## Future Enhancements

- Pivot table generation
- Support for multiple sheets
- Advanced column mapping interface
- Batch processing for multiple file pairs
- Performance optimization for large datasets
