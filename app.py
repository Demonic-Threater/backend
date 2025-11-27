from flask import Flask, request, send_file
from flask_cors import CORS
from docxtpl import DocxTemplate
from docx import Document
import tempfile
import json
import os
import subprocess

app = Flask(__name__)
CORS(app)

@app.route('/generate', methods=['POST'])
def generate():
    student_name = request.form.get('student_name', 'Student')
    class_name = request.form.get('class', '')
    registration_no = request.form.get('registration_no', '')
    roll_no = request.form.get('roll_no', '')
    start_year = request.form.get('start_year', '')
    end_year = request.form.get('end_year', '')
    subjects = json.loads(request.form.get('subjects', '[]'))

    template_path = os.path.join(os.path.dirname(__file__), "template.docx")

    merged_doc = Document()

    for idx, subject in enumerate(subjects):
        context = {
            "student_name": student_name,
            "subject": subject,
            "class": class_name,
            "registration_no": registration_no,
            "roll_no": roll_no,
            "start_year": start_year,
            "end_year": end_year
        }

        doc = DocxTemplate(template_path)
        doc.render(context)

        t = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(t.name)
        t.close()

        tmp_doc = Document(t.name)
        for element in tmp_doc.element.body:
            merged_doc.element.body.append(element)

        if idx < len(subjects) - 1:
            merged_doc.add_page_break()

        os.unlink(t.name)

    # Save merged DOCX
    out_doc = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    merged_doc.save(out_doc.name)
    out_doc.close()

    # Convert to PDF — libreoffice will create SAME FILENAME with .pdf
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", "/tmp",
        out_doc.name
    ], check=True)

    # Find the generated PDF
    pdf_path = out_doc.name.replace(".docx", ".pdf")

    # Remove DOCX
    os.unlink(out_doc.name)

    return send_file(pdf_path, as_attachment=True,
                     download_name=f"{student_name}_frontpage.pdf")

