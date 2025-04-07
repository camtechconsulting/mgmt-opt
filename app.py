
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from docx.shared import Inches
from datetime import datetime
import os
import tempfile
import mimetypes
import pdfplumber
import pytesseract
from PIL import Image
from pptx import Presentation
import pandas as pd
from openai import OpenAI
import io

app = Flask(__name__)
CORS(app)

REPORT_FOLDER = os.path.join(app.root_path, 'static', 'reports')
LOGO_PATH = os.path.join(app.root_path, 'static', 'logo.png')
os.makedirs(REPORT_FOLDER, exist_ok=True)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def extract_text(file):
    mime_type, _ = mimetypes.guess_type(file.filename)
    ext = os.path.splitext(file.filename)[-1].lower()

    if ext in [".docx"]:
        from docx import Document as DocxDocument
        doc = DocxDocument(file)
        return "\n".join([p.text for p in doc.paragraphs])

    elif ext in [".pdf"]:
        with pdfplumber.open(file) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)

    elif ext in [".png", ".jpg", ".jpeg"]:
        image = Image.open(file.stream)
        return pytesseract.image_to_string(image)

    elif ext in [".pptx"]:
        prs = Presentation(file)
        return "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text"))

    elif ext in [".xlsx"]:
        xls = pd.read_excel(file, sheet_name=None)
        return "\n".join(str(df) for df in xls.values())

    else:
        return f"Unsupported file type: {ext}"

def generate_section(prompt):
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            { "role": "system", "content": "You are an expert in business management and workflow optimization." },
            { "role": "user", "content": prompt }
        ],
        temperature=0.7
    )
    return response.choices[0].message.content.strip()

def add_logo(doc):
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    header = section.header
    paragraph = header.paragraphs[0]
    run = paragraph.add_run()
    if os.path.exists(LOGO_PATH):
        run.add_picture(LOGO_PATH, width=Inches(1.73), height=Inches(0.83))
        paragraph.alignment = 1

@app.route("/")
def home():
    return "Management Optimization Backend is Running!"

@app.route("/generate", methods=["POST"])
def generate_report():
    files = [request.files.get("doc1"), request.files.get("doc2"), request.files.get("doc3")]
    context = ""

    for file in files:
        if file and file.filename:
            context += extract_text(file) + "\n"

    if not context.strip():
        return jsonify({ "error": "No valid content extracted." }), 400

    sections = {
        "Executive Summary": "Provide an overview of the business's current management effectiveness, as observed in the uploaded materials.",
        "Leadership & Decision-Making": "Evaluate the leadership structure, delegation, and decision-making practices. Recommend improvements.",
        "Workflow Efficiency": "Analyze the workflows and organizational structure. Identify bottlenecks and inefficiencies.",
        "Team Performance & Culture": "Assess team dynamics, accountability systems, and management culture.",
        "Strategic Management Insights": "Offer long-term strategy insights, risk management, and leadership development ideas.",
        "Recommendations": "Summarize data-driven recommendations and propose a step-by-step optimization plan."
    }

    doc = Document()
    add_logo(doc)
    doc.add_heading("Management Optimization Report", 0)

    for heading, instruction in sections.items():
        doc.add_heading(heading, level=1)
        try:
            content = generate_section(f"{instruction}\n\nBusiness Context:\n{context}")
            doc.add_paragraph(content)
        except Exception as e:
            doc.add_paragraph(f"Error generating this section: {e}")

    filename = f"management_report_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    file_path = os.path.join(REPORT_FOLDER, filename)
    doc.save(file_path)

    return jsonify({ "download_url": f"/static/reports/{filename}" })

@app.route("/static/reports/<path:filename>")
def download_file(filename):
    return send_from_directory(REPORT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
