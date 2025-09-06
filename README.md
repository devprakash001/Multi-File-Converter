# Word to PDF Converter (Flask)

A simple Flask web application to convert Word (.docx) files to high-quality PDF files. Uses `docx2pdf` for conversion (Windows only).

## Features
- Upload `.docx` files and convert to PDF
- Download the converted PDF
- Clean Bootstrap UI
- Error handling for unsupported files and failed conversions

## Requirements
- Python 3.7+
- Windows OS (for `docx2pdf`)

## Installation

1. **Clone the repository:**
   ```bash
   git clone <repo-url>
   cd PDF
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the app:**
   ```bash
   python app.py
   ```

5. **Open your browser:**
   Go to [http://127.0.0.1:5000](http://127.0.0.1:5000)

## Notes
- Only `.docx` files are supported.
- Converted PDFs and uploads are stored temporarily in `converted/` and `uploads/` folders.
- For production, set a secure `app.secret_key` in `app.py`.

## License
MIT 