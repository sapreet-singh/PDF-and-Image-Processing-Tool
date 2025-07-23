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

## 📁 Supported Formats

### Input Formats
- **ZIP Archives**: Containing PNG, JPG, JPEG, BMP, TIFF, GIF images
- **PDF Files**: Both text-based and image-based (scanned) PDFs
- **Image Files**: PNG, JPG, JPEG, BMP, TIFF, GIF

### Output Formats
- **Excel Files**: Professionally formatted .xlsx with contact data
- **Text Files**: Raw extracted text for debugging
- **PDF Files**: Converted from ZIP archives (temporary)

## ⚙️ How the Tool Works - Step by Step

### **Phase 1: Input Processing**
1. **File Detection**: Tool identifies input file type (ZIP or PDF)
2. **Validation**: Checks file existence and format compatibility
3. **Extraction**: For ZIP files, extracts all images to the temporary directory

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

This project is open source. Please feel free to use and modify as you need.
