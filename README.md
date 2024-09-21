
**Description**:  
This Python application captures images from a webcam, processes them to extract text using OCR, and saves the extracted text to a Word document.

**Features**:
- **Capture Images**: Take pictures using a webcam and save them.
- **Convert to PDF**: Combine captured images into a single PDF.
- **OCR**: Extract text from PDF pages using Tesseract OCR.
- **Save to Word**: Output the extracted text to a Word document.

**Requirements**:
- Python 3
- OpenCV (`cv2`)
- PyMuPDF (`fitz`)
- Tesseract OCR
- pytesseract
- Python-docx (`docx`)
- PIL (Pillow)
- tkinter (included with Python)

**Setup**:
1. Install the required libraries:
   ```bash
   pip install opencv-python pymupdf pytesseract python-docx pillow
   ```

2. Ensure Tesseract OCR is installed and update the path in the script:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

**Usage**:
1. **Select Folder**: Choose a folder to save captured images.
2. **Select Output File**: Specify the location for the output Word document.
3. **Start Conversion**: Click the button to begin capturing images, converting to PDF, extracting text, and saving to Word.

**License**: Open-source. Free to use and modify.

**Youtube Video**: https://youtu.be/qVjHX9f56wI?si=N7yLEIsaCoEEeJsN
