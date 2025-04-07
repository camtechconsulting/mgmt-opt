
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from datetime import datetime
import os
import tempfile

app = Flask(__name__)
CORS(app)

REPORT_FOLDER = os.path.join(app.root_path, 'static', 'reports')
os.makedirs(REPORT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return "Management Optimization Backend is Live!"

@app.route('/generate', methods=['POST'])
def generate_report():
    files = [request.files.get('file1'), request.files.get('file2'), request.files.get('file3')]
    context = ""

    for file in files:
        if file:
            try:
                content = file.read().decode("utf-8", errors="ignore")
            except:
                content = "Unable to read file content."
            context += f"\n--- {file.filename} ---\n{content}\n"

    if not context.strip():
        return jsonify({"error": "No valid file content found."}), 400

    doc = Document()
    doc.add_heading("Management Optimization Report", 0)

    sections = [
        ("Executive Summary", "Provide a high-level overview of current management practices, leadership structure, and general observations."),
        ("1. Leadership & Team Structure", "Evaluate clarity of leadership roles, reporting structure, and overall org chart efficiency."),
        ("2. Workflow & Productivity", "Analyze task delegation, time management, and communication bottlenecks."),
        ("3. Decision-Making Processes", "Assess how decisions are made, who is involved, and effectiveness of those methods."),
        ("4. Management KPIs & Performance", "Review key metrics that track management effectiveness."),
        ("5. Recommendations & Optimization", "Offer actionable strategies to improve management efficiency, team performance, and scalability."),
        ("Conclusion", "Summarize the state of the business's management systems and suggest next steps.")
    ]

    for title, instruction in sections:
        doc.add_heading(title, level=1)
        doc.add_paragraph(f"(Section generated based on uploaded content and instruction: {instruction})")
        doc.add_paragraph("Generated content here...
")

    filename = f"management_report_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    file_path = os.path.join(REPORT_FOLDER, filename)
    doc.save(file_path)

    return jsonify({'download_url': f'/static/reports/{filename}'})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
