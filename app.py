import os
import zipfile
import uuid
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename
from docx2pdf import convert as docx2pdf_convert
import img2pdf
import pypandoc
from PIL import Image
import shutil

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {
    'word': {'doc', 'docx'},
    'powerpoint': {'ppt', 'pptx'},
    'excel': {'xls', 'xlsx'},
    'text': {'txt'},
    'image': {'jpg', 'jpeg', 'png'},
    'markdown': {'md'},
    'pdf': {'pdf'}
}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.secret_key = 'supersecretkey'  # Change this in production

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename, conversion_type=None):
    ext = filename.rsplit('.', 1)[-1].lower()
    if conversion_type and conversion_type in ALLOWED_EXTENSIONS:
        return ext in ALLOWED_EXTENSIONS[conversion_type]
    return any(ext in exts for exts in ALLOWED_EXTENSIONS.values())

def get_conversion_type(filename):
    ext = filename.rsplit('.', 1)[-1].lower()
    for ctype, exts in ALLOWED_EXTENSIONS.items():
        if ext in exts:
            return ctype
    return None

def convert_file(filepath, conversion_type, output_dir):
    ext = filepath.rsplit('.', 1)[-1].lower()
    basename = os.path.basename(filepath).rsplit('.', 1)[0]
    output_files = []
    if conversion_type == 'word':
        import pythoncom
        pythoncom.CoInitialize()
        pdf_path = os.path.join(output_dir, basename + '.pdf')
        docx2pdf_convert(filepath, pdf_path)
        pythoncom.CoUninitialize()
        output_files.append(pdf_path)
    elif conversion_type == 'powerpoint':
        # Use pypandoc or unoconv/LibreOffice for ppt/pptx
        pdf_path = os.path.join(output_dir, basename + '.pdf')
        pypandoc.convert_file(filepath, 'pdf', outputfile=pdf_path)
        output_files.append(pdf_path)
    elif conversion_type == 'excel':
        pdf_path = os.path.join(output_dir, basename + '.pdf')
        pypandoc.convert_file(filepath, 'pdf', outputfile=pdf_path)
        output_files.append(pdf_path)
    elif conversion_type == 'text':
        pdf_path = os.path.join(output_dir, basename + '.pdf')
        pypandoc.convert_file(filepath, 'pdf', outputfile=pdf_path)
        output_files.append(pdf_path)
    elif conversion_type == 'image':
        pdf_path = os.path.join(output_dir, basename + '.pdf')
        with open(pdf_path, "wb") as f:
            f.write(img2pdf.convert(filepath))
        output_files.append(pdf_path)
    elif conversion_type == 'markdown':
        pdf_path = os.path.join(output_dir, basename + '.pdf')
        pypandoc.convert_file(filepath, 'pdf', outputfile=pdf_path)
        output_files.append(pdf_path)
    elif conversion_type == 'pdf':
        # Optional: PDF to Word
        docx_path = os.path.join(output_dir, basename + '.docx')
        pypandoc.convert_file(filepath, 'docx', outputfile=docx_path)
        output_files.append(docx_path)
    return output_files

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('file')
        conversion_type = request.form.get('conversion_type')
        if not files or files[0].filename == '':
            flash('No file selected')
            return redirect(request.url)
        output_files = []
        temp_uploads = []
        for file in files:
            filename = secure_filename(file.filename)
            unique_id = str(uuid.uuid4())
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_id + '_' + filename)
            file.save(upload_path)
            temp_uploads.append(upload_path)
            # Detect conversion type if not selected
            ctype = conversion_type or get_conversion_type(filename)
            if not ctype or not allowed_file(filename, ctype):
                flash(f'Unsupported file type: {filename}')
                continue
            try:
                out_files = convert_file(upload_path, ctype, app.config['OUTPUT_FOLDER'])
                output_files.extend(out_files)
            except Exception as e:
                flash(f'Conversion failed for {filename}: {e}')
        # Clean up uploads
        for f in temp_uploads:
            try:
                os.remove(f)
            except Exception:
                pass
        if not output_files:
            return redirect(request.url)
        if len(output_files) == 1:
            download_link = url_for('download_file', filename=os.path.basename(output_files[0]))
            return render_template('index.html', download_link=download_link, success=True)
        else:
            # Zip multiple files
            zipname = f"converted_{uuid.uuid4().hex}.zip"
            zippath = os.path.join(app.config['OUTPUT_FOLDER'], zipname)
            with zipfile.ZipFile(zippath, 'w') as zipf:
                for f in output_files:
                    zipf.write(f, os.path.basename(f))
            download_link = url_for('download_file', filename=zipname)
            return render_template('index.html', download_link=download_link, success=True)
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True) 