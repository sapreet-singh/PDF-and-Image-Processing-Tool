# Contact Extraction Tool

🚀 **Intelligent Contact Information Extractor** - Convert ZIP files containing business card images or PDF documents into structured contact data with automatic Excel export.

## 🤔 What is this Tool?

This is an **automated contact information extraction system** that uses **Optical Character Recognition (OCR)** and **intelligent text parsing** to convert physical business cards, scanned documents, or PDF files into structured digital contact databases.

### **Key Capabilities:**
- 📸 **Image Processing**: Handles multiple image formats from ZIP archives
- 🔍 **OCR Technology**: Extracts text from images using Google's Tesseract engine
- 🧠 **Smart Parsing**: Uses AI-like algorithms to identify and categorize contact information
- 📊 **Data Structuring**: Organizes extracted data into professional Excel spreadsheets
- 🔧 **Batch Processing**: Processes multiple business cards or documents at once

## 💡 Why Build This Tool?

### **Problem Statement:**
- **Manual Data Entry**: Typing contact information from business cards is time-consuming and error-prone
- **Unstructured Data**: Business cards and documents contain valuable contact information in various formats
- **Scalability Issues**: Processing hundreds of business cards manually is impractical
- **Data Loss**: Physical business cards can be lost or damaged

### **Solution Benefits:**
- ⚡ **Time Saving**: Process hundreds of contacts in minutes instead of hours
- 🎯 **Accuracy**: Reduces human error in data entry
- 📈 **Scalability**: Handle large volumes of business cards efficiently
- 💾 **Digital Backup**: Convert physical cards to permanent digital records
- 🔄 **Standardization**: Consistent data format across all contacts
- 📱 **Integration Ready**: Excel output can be imported into CRM systems

## ✨ Features

✅ **ZIP to PDF Conversion** - Convert ZIP archives containing images into a single PDF document
✅ **Smart OCR Text Extraction** - Extract text from images and PDFs using Tesseract OCR
✅ **Intelligent Contact Parsing** - Automatically identify and extract contact information including:
- Names, job titles, and company names
- Email addresses and phone numbers (Mobile, Direct, HQ)
- Location and address information
✅ **Professional Excel Export** - Generate formatted Excel spreadsheets with extracted contacts
✅ **Multi-format Support** - Process both ZIP archives and PDF files
✅ **Text Cleaning & Normalization** - Remove OCR artifacts and standardize data

## 🚀 Quick Start

```bash
# Install Python dependencies
pip install pillow fpdf2 pytesseract PyMuPDF openpyxl

# Install Tesseract OCR (Windows)
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
# Or use chocolatey: choco install tesseract

# Process a ZIP file containing business card images
python main.py "business_cards.zip"

# Process a PDF document
python main.py "document.pdf" "my_contacts.xlsx"
```

## 📖 Usage

### Command Line Interface
```bash
# Basic usage - auto-generate output filename
python main.py input_file.zip

# Specify custom output Excel filename
python main.py input_file.pdf custom_contacts.xlsx
```

### Python API
```python
from main import process_file, convert_zip_to_pdf, extract_text_from_pdf

# Process any supported file type
excel_file = process_file("business_cards.zip", "contacts.xlsx")

# Individual functions
convert_zip_to_pdf("images.zip", "output.pdf")
extract_text_from_pdf("document.pdf", "extracted_text.txt")
```

## 🛠️ Complete Installation Guide

### **Prerequisites**
- **Python 3.7+** (recommended: Python 3.9 or higher)
- **Operating System**: Windows, macOS, or Linux
- **Internet Connection**: For downloading dependencies

### **Step 1: Clone or Download the Project**
```bash
# Option 1: Clone with Git
git clone <repository-url>
cd contact-extraction-tool

# Option 2: Download ZIP and extract
# Download the project files and navigate to the directory
```

### **Step 2: Python Dependencies Installation**

#### **Method 1: Using requirements.txt (Recommended)**
```bash
# Install all dependencies at once
pip install -r requirements.txt
```

#### **Method 2: Manual Installation**
```bash
# Install each package individually
pip install Pillow>=9.0.0
pip install fpdf2>=2.7.0
pip install pytesseract>=0.3.10
pip install PyMuPDF>=1.23.0
pip install openpyxl>=3.1.0
```

#### **Method 3: Using Virtual Environment (Best Practice)**
```bash
# Create virtual environment
python -m venv contact_extractor_env

# Activate virtual environment
# Windows:
contact_extractor_env\Scripts\activate
# macOS/Linux:
source contact_extractor_env/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### **Step 3: Tesseract OCR Installation**

#### **Windows Installation**
```bash
# Option 1: Download Installer
# 1. Visit: https://github.com/UB-Mannheim/tesseract/wiki
# 2. Download Windows installer
# 3. Run installer (default path: C:\Program Files\Tesseract-OCR\)
# 4. Add to PATH or update main.py with correct path

# Option 2: Using Chocolatey
choco install tesseract

# Option 3: Using Scoop
scoop install tesseract
```

#### **macOS Installation**
```bash
# Using Homebrew (recommended)
brew install tesseract

# Using MacPorts
sudo port install tesseract3
```

#### **Linux Installation**
```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install tesseract-ocr

# CentOS/RHEL/Fedora
sudo yum install tesseract
# or for newer versions:
sudo dnf install tesseract

# Arch Linux
sudo pacman -S tesseract
```

### **Step 4: Configuration**

#### **Tesseract Path Configuration**
If Tesseract is not automatically detected, update the path in `main.py`:

```python
# Windows (update if installed in different location)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# macOS (usually automatic, but if needed)
pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'

# Linux (usually automatic, but if needed)
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
```

### **Step 5: Verification**
```bash
# Test Python dependencies
python -c "import PIL, fpdf, pytesseract, fitz, openpyxl; print('All Python packages installed successfully!')"

# Test Tesseract installation
tesseract --version

# Test the tool
python main.py --help
```

## 📁 Supported Formats

### Input Formats
- **ZIP Archives**: Containing PNG, JPG, JPEG, BMP, TIFF, GIF images
- **PDF Files**: Both text-based and image-based (scanned) PDFs
- **Image Files**: PNG, JPG, JPEG, BMP, TIFF, GIF

### Output Formats
- **Excel Files**: Professional formatted .xlsx with contact data
- **Text Files**: Raw extracted text for debugging
- **PDF Files**: Converted from ZIP archives (temporary)

## 📚 Libraries Used and Why

### **Core Image Processing Libraries**

#### **Pillow (PIL) - Image Processing**
```python
from PIL import Image
```
- **Purpose**: Image manipulation, format conversion, and processing
- **Why Used**: Industry standard for Python image processing
- **Key Features**:
  - Supports all major image formats (PNG, JPG, JPEG, BMP, TIFF, GIF)
  - Image resizing and format conversion
  - Memory-efficient image handling
- **In This Tool**: Opens and processes images from ZIP archives

#### **FPDF2 - PDF Creation**
```python
from fpdf import FPDF
```
- **Purpose**: Lightweight PDF generation from images
- **Why Used**: Simple, fast, and reliable for basic PDF creation
- **Key Features**:
  - No external dependencies
  - Automatic page sizing
  - Image embedding with proper scaling
- **In This Tool**: Converts multiple images into a single PDF document

#### **PyMuPDF (fitz) - Advanced PDF Processing**
```python
import fitz  # PyMuPDF
```
- **Purpose**: Advanced PDF reading, text extraction, and image rendering
- **Why Used**: Most powerful Python PDF library with excellent performance
- **Key Features**:
  - Fast text extraction from PDF pages
  - Image rendering from PDF pages
  - Handles both text-based and scanned PDFs
- **In This Tool**: Extracts text and converts PDF pages to images for OCR

### **OCR (Optical Character Recognition)**

#### **pytesseract - OCR Engine**
```python
import pytesseract
```
- **Purpose**: Python wrapper for Google's Tesseract OCR engine
- **Why Used**: Most accurate and mature open-source OCR solution
- **Key Features**:
  - Supports 100+ languages
  - High accuracy text recognition
  - Configurable OCR parameters
- **In This Tool**: Converts images containing text into machine-readable text

### **Data Processing Libraries**

#### **openpyxl - Excel File Creation**
```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
```
- **Purpose**: Create and format Excel files (.xlsx)
- **Why Used**: Best Python library for modern Excel file creation
- **Key Features**:
  - Professional formatting and styling
  - Auto-sizing columns
  - Color schemes and fonts
- **In This Tool**: Creates formatted Excel spreadsheets with extracted contact data

### **Built-in Python Libraries**

#### **zipfile - Archive Processing**
```python
import zipfile
```
- **Purpose**: Extract and process ZIP archive files
- **Why Used**: Built-in Python library, no additional dependencies
- **In This Tool**: Extracts images from ZIP archives for processing

#### **re - Regular Expressions**
```python
import re
```
- **Purpose**: Pattern matching and text processing
- **Why Used**: Essential for parsing contact information patterns
- **In This Tool**: Extracts emails, phone numbers, and cleans OCR text

#### **os - File System Operations**
```python
import os
```
- **Purpose**: File and directory operations
- **Why Used**: Cross-platform file handling
- **In This Tool**: File existence checks, path operations, temporary file management

### **Library Selection Rationale**

1. **Performance**: All libraries are optimized for their specific tasks
2. **Reliability**: Mature, well-tested libraries with active communities
3. **Compatibility**: Cross-platform support for Windows, macOS, and Linux
4. **Minimal Dependencies**: Focused selection to avoid bloat
5. **Professional Output**: Libraries that produce high-quality, professional results

## ⚙️ How the Tool Works - Step by Step

### **Phase 1: Input Processing**
1. **File Detection**: Tool identifies input file type (ZIP or PDF)
2. **Validation**: Checks file existence and format compatibility
3. **Extraction**: For ZIP files, extracts all images to temporary directory

### **Phase 2: Image to PDF Conversion** (ZIP files only)
1. **Image Sorting**: Sorts images alphabetically for consistent processing
2. **Format Validation**: Filters supported image formats (PNG, JPG, JPEG, BMP, TIFF, GIF)
3. **PDF Creation**: Converts images to PDF with proper sizing and layout
4. **Cleanup**: Removes temporary files after processing

### **Phase 3: OCR Text Extraction**
1. **PDF Analysis**: Checks if PDF contains extractable text or requires OCR
2. **Text Extraction**: For text-based PDFs, extracts existing text directly
3. **OCR Processing**: For image-based PDFs, converts pages to images and applies Tesseract OCR
4. **Page Segmentation**: Processes each page separately for better accuracy

### **Phase 4: Intelligent Contact Parsing**
1. **Text Preprocessing**: Removes OCR artifacts and normalizes text
2. **Pattern Recognition**: Uses regex patterns to identify:
   - **Email Addresses**: Standard email format validation
   - **Phone Numbers**: Multiple formats (Mobile, Direct, HQ)
   - **Names**: Distinguishes person names from company names
   - **Job Titles**: Keyword-based title recognition
   - **Companies**: Business entity identification
   - **Locations**: Address and location parsing

### **Phase 5: Data Structuring & Export**
1. **Contact Assembly**: Organizes extracted information into structured records
2. **Data Validation**: Ensures data quality and completeness
3. **Excel Generation**: Creates formatted spreadsheet with:
   - Professional styling and headers
   - Auto-sized columns
   - Color-coded headers
4. **File Output**: Saves both raw text and structured Excel files

## 📊 Output Structure

The generated Excel file contains the following columns:
- **Name**: Person's full name
- **Title**: Job title or position
- **Company**: Company or organization name
- **Email**: Email address
- **Mobile Phone**: Mobile/cell phone number
- **Direct Phone**: Direct office line
- **HQ Phone**: Headquarters phone number
- **Location**: Address or location information

## 🔧 Configuration

### Tesseract Path Configuration
If Tesseract is not in your system PATH, update the path in `main.py`:
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

### Customizing Contact Extraction
You can modify the extraction patterns in `main.py`:
- **Title Keywords**: Add job titles to recognize in `title_keywords` list
- **Company Indicators**: Add company name patterns in `company_indicators` list
- **Phone Patterns**: Modify regex patterns for different phone number formats
- **Location Parsing**: Adjust location detection logic

## 🚨 Troubleshooting

### Common Issues

**"Tesseract not found" Error:**
- Ensure Tesseract is installed and in PATH
- Update `tesseract_cmd` path in `main.py`
- Restart terminal/IDE after installation

**Poor OCR Results:**
- Ensure images are high resolution and clear
- Check image orientation (rotate if needed)
- Verify good contrast between text and background

**Missing Contact Information:**
- Review the extracted text file to see raw OCR output
- Adjust extraction patterns for your specific document format
- Consider preprocessing images for better OCR results

## 📝 Example Output

```
📦 Processing ZIP file: business_cards.zip
Step 1: Converting ZIP to PDF...
Step 2: Extracting text from PDF...
✓ Text saved: business_cards_text.txt
Step 2: Extracting contact information...
✓ Found 15 contacts
Step 3: Creating Excel file...
✓ Excel file created: business_cards_contacts.xlsx

📊 Summary:
   • Input file: business_cards.zip
   • Text file: business_cards_text.txt
   • Excel file: business_cards_contacts.xlsx
   • Contacts extracted: 15
```

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve the tool!

## 📄 License

This project is open source. Feel free to use and modify as needed.
