from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from pdf2docx import Converter
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape
import pdfplumber
import pandas as pd
import pytesseract
import subprocess
import os
import zipfile
import time
import uuid
import os

app = Flask(__name__)

# UPLOAD_FOLDER = 'uploads'
# OUTPUT_FOLDER = 'outputs'

UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ---------------- CLEANUP ----------------
def clear_folder(folder):
    for f in os.listdir(folder):
        path = os.path.join(folder, f)
        if os.path.isfile(path):
            os.remove(path)

clear_folder(UPLOAD_FOLDER)
clear_folder(OUTPUT_FOLDER)


# ---------------- VALIDATION ----------------
def validate_files(action, files):
    if not action:
        return "No action selected"

    for file in files:
        ext = file.filename.rsplit('.', 1)[1].lower()

        if action == 'img2pdf' and ext not in ['jpg', 'jpeg', 'png']:
            return "Only images allowed"

        if action == 'pdf2docx' and ext != 'pdf':
            return "PDF required"

        if action == 'docx2pdf' and ext != 'docx':
            return "DOCX required"

        if action in ['mergepdf', 'splitpdf', 'pdf2excel', 'pdf2text', 'protectpdf', 'compresspdf'] and ext != 'pdf':
            return "PDF required"

        if action == 'excel2pdf' and ext not in ['xlsx', 'xls', 'csv']:
            return "Excel or CSV required"

    return None


# ---------------- FEATURES ----------------

# IMAGE → PDF
def img_to_pdf(image_paths, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    page_w, page_h = A4

    for img_path in image_paths:
        img = Image.open(img_path).convert('RGB')
        img_reader = ImageReader(img)

        img_w, img_h = img.size
        scale = min(page_w / img_w, page_h / img_h) * 0.9

        new_w = img_w * scale
        new_h = img_h * scale

        x = (page_w - new_w) / 2
        y = (page_h - new_h) / 2

        c.setFillColorRGB(1, 1, 1)
        c.rect(0, 0, page_w, page_h, fill=1)
        c.drawImage(img_reader, x, y, width=new_w, height=new_h)
        c.showPage()

    c.save()


# PDF → DOCX (SMART)
def pdf_to_docx(input_path, output_path):
    try:
        cv = Converter(input_path)
        cv.convert(output_path, multi_processing=True)
        cv.close()

        if os.path.exists(output_path) and os.path.getsize(output_path) > 2000:
            return

    except Exception as e:
        print("pdf2docx failed:", e)

    try:
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to", "docx",
            "--outdir", os.path.dirname(output_path),
            input_path
        ], check=True)

    except Exception as e:
        print("LibreOffice failed:", e)
        raise Exception("Conversion failed")


# DOCX → PDF
# def docx_to_pdf(input_path, output_path):
#     try:
#         subprocess.run([
#             "libreoffice",
#             "--headless",
#             "--convert-to", "pdf",
#             "--outdir", os.path.dirname(output_path),
#             input_path
#         ], check=True)

#     except Exception as e:
#         print("DOCX to PDF failed:", e)
#         raise Exception("Conversion failed")
    
def docx_to_pdf(input_path, output_path):
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", os.path.dirname(output_path),
        input_path
    ], check=True)

    # FIX: Rename to expected output
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    generated = os.path.join(os.path.dirname(output_path), base_name + ".pdf")

    if generated != output_path:
        os.rename(generated, output_path)


# MERGE PDF
def merge_pdfs(pdf_list, output_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()


# PDF → EXCEL
def pdf_to_excel(pdf_path, output_path):
    tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()

            if table and len(table) > 1:
                df = pd.DataFrame(table[1:], columns=table[0])
                tables.append(df)

    if tables:
        final = pd.concat(tables)
        final.to_excel(output_path, index=False)
    else:
        raise Exception("No tables found")


# EXCEL → PDF
# def excel_to_pdf(excel_path, output_path):
#     from reportlab.platypus import SimpleDocTemplate, Table

#     df = pd.read_excel(excel_path)
#     doc = SimpleDocTemplate(output_path)

#     data = [df.columns.tolist()] + df.values.tolist()
#     table = Table(data[:50])

#     doc.build([table])

def excel_to_pdf(excel_path, output_path, orientation='portrait', fit='fit'):
    styles = getSampleStyleSheet()

    # FIX: SAFE EXCEL READ
    try:
        ext = excel_path.split('.')[-1].lower()

        if ext == 'xlsx':
            df = pd.read_excel(excel_path, engine='openpyxl')

        elif ext == 'xls':
            df = pd.read_excel(excel_path, engine='xlrd')

        elif ext == 'csv':
            df = pd.read_csv(excel_path)

        else:
            raise Exception("Unsupported file format")

    except Exception as e:
        print("Read Error:", str(e))
        raise Exception("File corrupted or unsupported")
    
    if df.empty:
        raise Exception("File is empty")

    df = df.dropna(how='all')        # remove empty rows
    df = df.dropna(axis=1, how='all')  # remove empty columns

    # limit to avoid PDF crash
    df = df.head(100)
    df = df.iloc[:, :10]

    # SET ORIENTATION
    if orientation == 'landscape':
        page_size = landscape(A4)
    else:
        page_size = A4

    doc = SimpleDocTemplate(output_path, pagesize=page_size)

    data = [df.columns.tolist()]

    for row in df.values:
        wrapped_row = [Paragraph(str(cell), styles['Normal']) for cell in row]
        data.append(wrapped_row)

    table = Table(data)

    # AUTO WIDTH
    if fit == 'fit':
        total_width = page_size[0] - 40
        col_count = len(df.columns)
        col_width = total_width / col_count
        table._argW = [col_width] * col_count

    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))

    doc.build([table])


# SPLIT PDF
# def split_pdf_pages(input_path, pages, output_path):
#     reader = PdfReader(input_path)
#     writer = PdfWriter()

#     for p in pages:
#         if 0 <= p < len(reader.pages):
#             writer.add_page(reader.pages[p])

#     with open(output_path, "wb") as f:
#         writer.write(f)

def parse_page_range(page_range, total_pages):
    pages = set()

    parts = page_range.split(',')

    for part in parts:
        part = part.strip()

        if '-' in part:
            start, end = part.split('-')
            start = int(start)
            end = int(end)

            for i in range(start, end + 1):
                if 1 <= i <= total_pages:
                    pages.add(i - 1)  # convert to 0-based index

        else:
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p - 1)

    return sorted(pages)

def split_pdf_pages(input_path, page_range, output_path):
    reader = PdfReader(input_path)
    writer = PdfWriter()

    pages = parse_page_range(page_range, len(reader.pages))

    if not pages:
        raise Exception("Invalid page range")

    for p in pages:
        writer.add_page(reader.pages[p])

    with open(output_path, "wb") as f:
        writer.write(f)


# PDF → TEXT
def pdf_to_text(pdf_path, output_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""

    with open(output_path, "w") as f:
        f.write(text)


# OCR
def ocr_image(image_path, output_path):
    text = pytesseract.image_to_string(Image.open(image_path))
    with open(output_path, "w") as f:
        f.write(text)


# PROTECT PDF
# def protect_pdf(input_path, output_path):
#     reader = PdfReader(input_path)
#     writer = PdfWriter()

#     for page in reader.pages:
#         writer.add_page(page)

#     writer.encrypt("1234")

#     with open(output_path, "wb") as f:
#         writer.write(f)

def protect_pdf(input_path, output_path, password):
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(password)

    with open(output_path, "wb") as f:
        writer.write(f)


# COMPRESS PDF
def compress_pdf(input_path, output_path):
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)


# ---------------- ROUTES ----------------

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('files')
    print("Uploaded filename:", files[0].filename)  
    action = request.form.get('action')

    if not files or files[0].filename == '':
        return jsonify({'error': 'No file uploaded'}), 400

    error = validate_files(action, files)
    if error:
        return jsonify({'error': error}), 400

    saved_paths = []
    original_name = os.path.splitext(files[0].filename)[0]

    for file in files:
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)
        saved_paths.append(path)

    try:
        #output = os.path.join(OUTPUT_FOLDER, original_name + "_output")
        #output = os.path.join(OUTPUT_FOLDER, original_name)
        timestamp = int(time.time())

        safe_name = f"{original_name}_{action}_{timestamp}"
        output = os.path.join(OUTPUT_FOLDER, safe_name)

        # ACTIONS
        if action == 'img2pdf':
            output += ".pdf"
            img_to_pdf(saved_paths, output)

        elif action == 'pdf2docx':
            output += ".docx"
            pdf_to_docx(saved_paths[0], output)

        elif action == 'docx2pdf':
            output += ".pdf"
            docx_to_pdf(saved_paths[0], output)

        elif action == 'mergepdf':
            output = os.path.join(OUTPUT_FOLDER, "merged.pdf")
            merge_pdfs(saved_paths, output)

        elif action == 'pdf2excel':
            mode = request.form.get('mode', 'table')
            output += ".xlsx"

            if mode == 'table':
                pdf_to_excel(saved_paths[0], output)
            else:
                # fallback: extract full text into excel
                text_output = output.replace(".xlsx", ".txt")
                pdf_to_text(saved_paths[0], text_output)

                df = pd.DataFrame({"Text": open(text_output).read().split('\n')})
                df.to_excel(output, index=False)

        elif action == 'excel2pdf':
            output += ".pdf"

            orientation = request.form.get('orientation', 'portrait')
            fit = request.form.get('fit', 'fit')

            excel_to_pdf(saved_paths[0], output, orientation, fit)

        elif action == 'splitpdf':
            page_range = request.form.get('page_range')

            if not page_range:
                return jsonify({'error': 'Page range required'}), 400

            output += ".pdf"
            split_pdf_pages(saved_paths[0], page_range, output)

        elif action == 'pdf2text':
            output += ".txt"
            pdf_to_text(saved_paths[0], output)

        elif action == 'ocr':
            output += ".txt"
            ocr_image(saved_paths[0], output)

        elif action == 'protectpdf':
            password = request.form.get('password')

            if not password:
                return jsonify({'error': 'Password required'}), 400

            output += ".pdf"
            protect_pdf(saved_paths[0], output, password)

        elif action == 'compresspdf':
            output += ".pdf"
            compress_pdf(saved_paths[0], output)

        else:
            return jsonify({'error': 'Invalid action'}), 400

        # CLEANUP AFTER DOWNLOAD
        @after_this_request
        def cleanup(response):
            try:
                for path in saved_paths:
                    if os.path.exists(path):
                        os.remove(path)

                if os.path.exists(output):
                    os.remove(output)

            except Exception as e:
                print("Cleanup error:", e)

            return response

       # return send_file(output, as_attachment=True)
        #return send_file( output, as_attachment=True, download_name=os.path.basename(output))\
        download_name = f"{original_name}_{action}_{uuid.uuid4().hex[:6]}{os.path.splitext(output)[1]}"

        return send_file(
            output,
            as_attachment=True,
            download_name=download_name
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


# if __name__ == '__main__':
#     app.run(debug=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)