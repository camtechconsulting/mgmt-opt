
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from docx.shared import Inches
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

REPORT_FOLDER = os.path.join(app.root_path, 'static', 'reports')
LOGO_PATH = os.path.join(app.root_path, 'static', 'logo.png')
os.makedirs(REPORT_FOLDER, exist_ok=True)

def add_logo(doc):
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    header = section.header
    paragraph = header.paragraphs[0]
    run = paragraph.add_run()
    if os.path.exists(LOGO_PATH):
        run.add_picture(LOGO_PATH, width=Inches(1.73), height=Inches(0.83))
        paragraph.alignment = 1

@app.route('/')
def home():
    return "Basic Save Test Backend Running!"

@app.route('/generate', methods=['POST'])
def generate_report():
    print(f"✅ Checking save path: {REPORT_FOLDER}")
    if not os.path.exists(REPORT_FOLDER):
        print("❌ REPORT_FOLDER does not exist!")
    else:
        print("✅ REPORT_FOLDER exists and is accessible.")

    doc = Document()
    add_logo(doc)
    doc.add_heading("Management Optimization Test Report", 0)
    doc.add_paragraph("This is a test report with debug logging enabled.")

    filename = f"test_report_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    file_path = os.path.join(REPORT_FOLDER, filename)
    print(f"📄 Intended file path: {file_path}")

    try:
        doc.save(file_path)
        print(f"✅ Successfully saved report at: {file_path}")
    except Exception as e:
        print(f"❌ Error saving report: {e}")

    return jsonify({'download_url': f'/static/reports/{filename}'})

@app.route('/static/reports/<path:filename>')
def download_file(filename):
    return send_from_directory(REPORT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
