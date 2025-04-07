from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
from datetime import datetime
import os

app = Flask(__name__, static_url_path='/static')
CORS(app)

REPORT_FOLDER = os.path.join(app.root_path, 'static', 'reports')
os.makedirs(REPORT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return "Management Optimization Backend is Running!"

@app.route('/generate', methods=['POST'])
def generate_report():
    try:
        files = request.files.getlist("documents")
        if not files:
            return jsonify({'error': 'No files uploaded'}), 400

        context = ""
        for file in files:
            context += f"--- {file.filename} ---\n"
            context += file.read().decode(errors='ignore') + "\n"

        # Simulated analysis for demonstration purposes
        doc = Document()
        doc.add_heading("Management Optimization Report", 0)
        doc.add_paragraph("This report analyzes uploaded management documents to identify inefficiencies, patterns, and provides recommendations.\n")
        doc.add_paragraph("Raw Extracted Context:")
        doc.add_paragraph(context)

        filename = f"management_report_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
        file_path = os.path.join(REPORT_FOLDER, filename)
        doc.save(file_path)

        return jsonify({'download_url': f'/static/reports/{filename}'})
    except Exception as e:
        return jsonify({'error': f"Failed to generate report: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
