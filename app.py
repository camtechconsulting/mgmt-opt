
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from docx.shared import Inches
from datetime import datetime
import fitz
import pandas as pd
import openai
import os
import re
from pptx import Presentation
from PIL import Image

app = Flask(__name__)
CORS(app)

openai.api_key = os.getenv("OPENAI_API_KEY")
REPORT_FOLDER = os.path.join(app.root_path, 'static', 'reports')
LOGO_PATH = os.path.join(app.root_path, 'static', 'logo.png')
os.makedirs(REPORT_FOLDER, exist_ok=True)

def extract_text_docx(file):
    from docx import Document as DocxDocument
    try:
        doc = DocxDocument(file)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception:
        return ""

def extract_text_pdf(file):
    try:
        pdf = fitz.open(stream=file.read(), filetype="pdf")
        return "\n".join([page.get_text() for page in pdf])
    except Exception:
        return ""

def extract_text_pptx(file):
    try:
        prs = Presentation(file)
        text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text.append(shape.text)
        return "\n".join(text)
    except Exception:
        return ""

def extract_text_excel(file):
    try:
        df = pd.read_excel(file)
        return df.to_string(index=False)
    except Exception:
        return ""

def extract_text_image(file):
    try:
        image = Image.open(file.stream)
        return image.filename
    except Exception:
        return ""

def extract_text(file_storage):
    filename = file_storage.filename.lower()
    if filename.endswith(".docx"):
        return extract_text_docx(file_storage)
    elif filename.endswith(".pdf"):
        return extract_text_pdf(file_storage)
    elif filename.endswith(".pptx"):
        return extract_text_pptx(file_storage)
    elif filename.endswith(".xlsx"):
        return extract_text_excel(file_storage)
    elif filename.endswith((".png", ".jpg", ".jpeg")):
        return extract_text_image(file_storage)
    else:
        return ""

def clean_markdown(text):
    text = re.sub(r'^#+\s*', '', text, flags=re.MULTILINE)
    return text.replace("*", "").strip()

def extract_table_data(text):
    table = []
    lines = text.strip().splitlines()
    for line in lines:
        if '|' in line:
            row = [cell.strip() for cell in line.split('|') if cell.strip()]
            if row:
                table.append(row)
    return table if len(table) >= 2 else None

def generate_section(prompt):
    try:
        print("Calling OpenAI API...")
        response = openai.ChatCompletion.create(
            model="gpt-4-0125-preview",
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional business consultant. Generate deep management-level insights and action plans with recommendations and tables where applicable."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.7,
            max_tokens=2000
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        print("OpenAI API error:", e)
        return "Error generating this section."

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
    doc1 = request.files.get("doc1")
    doc2 = request.files.get("doc2")
    doc3 = request.files.get("doc3")

    context = ""
    if doc1: context += extract_text(doc1) + "\n"
    if doc2: context += extract_text(doc2) + "\n"
    if doc3: context += extract_text(doc3) + "\n"

    if not context.strip():
        return jsonify({"error": "No valid input provided."}), 400

    doc = Document()
    add_logo(doc)
    doc.add_heading("Management Metric Optimization Report", 0)

    sections = [
        ("Executive Summary", "Summarize the current state of management, highlighting both strengths and weaknesses."),
        ("Leadership Assessment", "Evaluate leadership communication, structure, and effectiveness."),
        ("Workflow Analysis", "Assess efficiency, time usage, communication pipelines, and processes."),
        ("Team Performance", "Analyze employee performance, engagement, and task distribution."),
        ("Strategic Planning", "Suggest improvements to decision-making and long-term planning."),
        ("Risk & Recommendations", "Identify managerial risks and how to mitigate them."),
    ]

    for title, instruction in sections:
        doc.add_heading(title, level=1)
        prompt = f"{instruction}\n\nBusiness Context:\n{context}"
        result = generate_section(prompt)
        table_data = extract_table_data(result)
        if table_data:
            table = doc.add_table(rows=1, cols=len(table_data[0]))
            table.style = "Table Grid"
            hdr_cells = table.rows[0].cells
            for i, val in enumerate(table_data[0]):
                hdr_cells[i].text = val
            for row_data in table_data[1:]:
                row_cells = table.add_row().cells
                for i, val in enumerate(row_data):
                    if i < len(row_cells):
                        row_cells[i].text = val
        else:
            doc.add_paragraph(clean_markdown(result))

    filename = f"management_report_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    file_path = os.path.join(REPORT_FOLDER, filename)

    try:
        doc.save(file_path)
        print(f"✅ Saved at {file_path}")
    except Exception as e:
        print(f"❌ Failed to save: {e}")

    return jsonify({"download_url": f"/static/reports/{filename}"})

@app.route('/static/reports/<path:filename>')
def download_file(filename):
    return send_from_directory(REPORT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
