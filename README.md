# 📄 FileForge - Flask File Converter

A powerful and user-friendly web application built using Flask that allows users to perform multiple file conversion and PDF manipulation tasks in one place.

💼 Built as part of my full-stack development portfolio to demonstrate real-world backend processing and file handling.

---

## 🚀 Features

- 🖼 Image to PDF conversion
- 📄 PDF to DOCX conversion
- 🔁 Merge multiple PDF files
- ✂️ Split PDF by custom page range
- 📊 PDF to Excel (table extraction)
- 📉 Excel/CSV to PDF
- 📝 PDF to Text extraction
- 🔍 OCR (Image to Text)
- 🔒 Password protect PDF files
- 📦 Compress PDF files

---

## 🛠 Tech Stack

- **Backend:** Python (Flask)
- **Frontend:** HTML, CSS, JavaScript
- **Libraries:**
  - PyPDF2
  - pdfplumber
  - pandas
  - reportlab
  - pytesseract
  - pdf2docx
  - Pillow

---

## ⚙️ How It Works

1. User uploads file(s)
2. Selects desired operation
3. Server processes the file
4. File is downloaded instantly
5. Files are automatically deleted after processing

---

## ▶️ Run Locally

```bash
pip install -r requirements.txt
python app.py