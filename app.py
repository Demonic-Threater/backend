from flask import Flask, request, send_file
from flask_cors import CORS
from docxtpl import DocxTemplate
from docx import Document
from docx2pdf import convert
import tempfile
import json
import os

app = Flask(__name__)
CORS(app)  # Allow requests from frontend

@app.route('/generate', methods=['POST'])
def generate():
    # Extract form data
    student_name = request.form.get('student_name', 'Student')
    class_name = request.form.get('class', '')
    registration_no = request.form.get('registration_no', '')
    roll_no = request.form.get('roll_no', '')
    start_year = request.form.get('start_year', '')
    end_year = request.form.get('end_year', '')
    subjects = json.loads(request.form.get('subjects', '[]'))

    template_path = os.path.join(os.path.dirname(__file__), "template.docx")

    merged_doc = Document()  # Final document

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

        # Save temp page
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp_file.name)
        tmp_file.close()

        # Append to merged_doc
        tmp_doc = Document(tmp_file.name)
        for element in tmp_doc.element.body:
            merged_doc.element.body.append(element)

        # Add page break except last page
        if idx < len(subjects) - 1:
            merged_doc.add_page_break()

        os.unlink(tmp_file.name)

    # Save merged doc to temp
    temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    merged_doc.save(temp_docx.name)
    temp_docx.close()

    # Convert to PDF
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    convert(temp_docx.name, temp_pdf.name)
    os.unlink(temp_docx.name)

    return send_file(temp_pdf.name, as_attachment=True,
                     download_name=f"{student_name}_frontpage.pdf")
