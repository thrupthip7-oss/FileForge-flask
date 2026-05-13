from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from pdf2docx import Converter
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import pdfplumber
import pandas as pd
import pytesseract
import subprocess
import os
import time
import uuid
import fitz
import json

app = Flask(__name__)

# Using /tmp for serverless compatibility (like Vercel/Render)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------------- CLEANUP ON START ----------------
def clear_folder(folder):
    for f in os.listdir(folder):
        path = os.path.join(folder, f)
        if os.path.isfile(path):
            try:
                os.remove(path)
            except:
                pass

clear_folder(UPLOAD_FOLDER)
clear_folder(OUTPUT_FOLDER)

# ---------------- VALIDATION ----------------
def validate_files(action, files):
    if not action:
        return "No action selected"
    
    # Define which actions ARE allowed to have multiple files
    multi_file_actions = ['img2pdf', 'mergepdf']
    
    if len(files) > 1 and action not in multi_file_actions:
        return f"Only one file allowed for {action}"

    for file in files:
        if not file.filename:
            return "One of the uploaded files has no name"
            
        ext = file.filename.rsplit('.', 1)[1].lower()

        if action == 'img2pdf' and ext not in ['jpg', 'jpeg', 'png']:
            return "Only images (jpg, png) allowed for PDF conversion"

        if action in ['pdf2docx', 'mergepdf', 'splitpdf', 'pdf2excel', 'pdf2text', 'protectpdf', 'compresspdf','whiteout'] and ext != 'pdf':
            return "PDF file required"

        if action == 'docx2pdf' and ext != 'docx':
            return "DOCX file required"

        if action == 'excel2pdf' and ext not in ['xlsx', 'xls', 'csv']:
            return "Excel or CSV required"

    return None

# ---------------- CORE FEATURES ----------------

def img_to_pdf(image_paths, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    page_w, page_h = A4
    for img_path in image_paths:
        img = Image.open(img_path).convert('RGB')
        img_reader = ImageReader(img)
        img_w, img_h = img.size
        scale = min(page_w / img_w, page_h / img_h) * 0.9
        new_w, new_h = img_w * scale, img_h * scale
        x, y = (page_w - new_w) / 2, (page_h - new_h) / 2
        c.setFillColorRGB(1, 1, 1)
        c.rect(0, 0, page_w, page_h, fill=1)
        c.drawImage(img_reader, x, y, width=new_w, height=new_h)
        c.showPage()
    c.save()

def pdf_to_docx(input_path, output_path):
    cv = Converter(input_path)
    cv.convert(output_path, multi_processing=False)
    cv.close()
    return output_path

def docx_to_pdf(input_path, output_path):
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf",
        "--outdir", os.path.dirname(output_path), input_path
    ], check=True)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    generated = os.path.join(os.path.dirname(output_path), base_name + ".pdf")
    if os.path.exists(generated) and generated != output_path:
        os.rename(generated, output_path)

def merge_pdfs(pdf_list, output_path):
    merger = PdfMerger()
    # Log it to your terminal to see the order the server receives
    print(f"Merging files in this order: {pdf_list}") 
    
    for pdf in pdf_list:
        merger.append(pdf)
    
    merger.write(output_path)
    merger.close()

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

def excel_to_pdf(excel_path, output_path, orientation='portrait', fit='fit'):
    styles = getSampleStyleSheet()
    ext = excel_path.split('.')[-1].lower()
    if ext == 'xlsx': df = pd.read_excel(excel_path, engine='openpyxl')
    elif ext == 'xls': df = pd.read_excel(excel_path, engine='xlrd')
    elif ext == 'csv': df = pd.read_csv(excel_path)
    else: raise Exception("Unsupported file format")
    
    if df.empty: raise Exception("File is empty")
    df = df.dropna(how='all').dropna(axis=1, how='all').head(100).iloc[:, :10]

    page_size = landscape(A4) if orientation == 'landscape' else A4
    doc = SimpleDocTemplate(output_path, pagesize=page_size)
    data = [df.columns.tolist()]
    for row in df.values:
        data.append([Paragraph(str(cell), styles['Normal']) for cell in row])

    table = Table(data)
    if fit == 'fit':
        col_width = (page_size[0] - 40) / len(df.columns)
        table._argW = [col_width] * len(df.columns)

    table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('FONTSIZE', (0, 0), (-1, -1), 8)]))
    doc.build([table])

def parse_page_range(page_range, total_pages):
    pages = set()
    for part in page_range.split(','):
        part = part.strip()
        if '-' in part:
            start, end = map(int, part.split('-'))
            for i in range(start, end + 1):
                if 1 <= i <= total_pages: pages.add(i - 1)
        else:
            p = int(part)
            if 1 <= p <= total_pages: pages.add(p - 1)
    return sorted(pages)

def split_pdf_pages(input_path, page_range, output_path):
    reader = PdfReader(input_path)
    pages = parse_page_range(page_range, len(reader.pages))
    if not pages: raise Exception("Invalid page range")
    writer = PdfWriter()
    for p in pages:
        writer.add_page(reader.pages[p])
    with open(output_path, "wb") as f:
        writer.write(f)

def pdf_to_text(pdf_path, output_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

def ocr_image(image_path, output_path):
    text = pytesseract.image_to_string(Image.open(image_path))
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

def protect_pdf(input_path, output_path, password):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(password)
    with open(output_path, "wb") as f:
        writer.write(f)

def compress_pdf(input_path, output_path):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)

def whiteout_pdf(input_path, output_path, whiteout_areas):

    doc = fitz.open(input_path)

    for area in whiteout_areas:

        page_index = area['page'] - 1

        if page_index < 0 or page_index >= len(doc):
            continue

        page = doc[page_index]

        rect = fitz.Rect(
            area['x'],
            area['y'],
            area['x'] + area['width'],
            area['y'] + area['height']
        )

        page.add_redact_annot(
            rect,
            fill=(1, 1, 1)
        )

        page.apply_redactions()

    doc.save(
        output_path,
        garbage=4,
        deflate=True
    )

    doc.close()

# ---------------- ROUTES ----------------

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('files')
    action = request.form.get('action')

    if not files or files[0].filename == '':
        return jsonify({'error': 'No file uploaded'}), 400

    error = validate_files(action, files)
    if error:
        return jsonify({'error': error}), 400

    saved_paths = []
    original_name = os.path.splitext(files[0].filename)[0]
    
    # Save input files
    for file in files:
        path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}_{file.filename}")
        file.save(path)
        saved_paths.append(path)

    try:
        timestamp = int(time.time())
        output_filename = f"{original_name}_{action}_{timestamp}"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        # Logic Branching
        if action == 'img2pdf':
            output_path += ".pdf"
            img_to_pdf(saved_paths, output_path)
        elif action == 'pdf2docx':
            output_path += ".docx"
            pdf_to_docx(saved_paths[0], output_path)
        elif action == 'docx2pdf':
            output_path += ".pdf"
            docx_to_pdf(saved_paths[0], output_path)
        elif action == 'mergepdf':
            output_path = os.path.join(OUTPUT_FOLDER, f"merged_{timestamp}.pdf")
            merge_pdfs(saved_paths, output_path)
        elif action == 'pdf2excel':
            output_path += ".xlsx"
            pdf_to_excel(saved_paths[0], output_path)
        elif action == 'excel2pdf':
            output_path += ".pdf"
            excel_to_pdf(saved_paths[0], output_path, request.form.get('orientation'), request.form.get('fit'))
        elif action == 'splitpdf':
            output_path += ".pdf"
            split_pdf_pages(saved_paths[0], request.form.get('page_range'), output_path)
        elif action == 'pdf2text':
            output_path += ".txt"
            pdf_to_text(saved_paths[0], output_path)
        elif action == 'ocr':
            output_path += ".txt"
            ocr_image(saved_paths[0], output_path)
        elif action == 'protectpdf':
            output_path += ".pdf"
            protect_pdf(saved_paths[0], output_path, request.form.get('password'))
        elif action == 'compresspdf':
            output_path += ".pdf"
            compress_pdf(saved_paths[0], output_path)
        elif action == 'whiteout':
            output_path += ".pdf"

            whiteout_areas = request.form.get('whiteout_areas')

            if not whiteout_areas:
                raise Exception("No whiteout areas provided")

            import json

            whiteout_areas = json.loads(whiteout_areas)

            whiteout_pdf(
                saved_paths[0],
                output_path,
                whiteout_areas
            )
        else:
            return jsonify({'error': 'Invalid action'}), 400

        @after_this_request
        def cleanup(response):
            for p in saved_paths:
                if os.path.exists(p): os.remove(p)
            # Note: We don't delete output_path here because send_file needs it
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name=os.path.basename(output_path)
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)