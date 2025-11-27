from flask import Flask, request, send_file
from docxtpl import DocxTemplate
from docx import Document
from flask_cors import CORS
import tempfile
import json
import os
import subprocess

app = Flask(__name__)
CORS(app)  # Allow all origins

# Function to convert DOCX -> PDF using LibreOffice headless
def convert_docx_to_pdf(input_docx):
    output_dir = tempfile.mkdtemp()
    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        input_docx,
        "--outdir", output_dir
    ], check=True)
    pdf_file = os.path.join(output_dir, os.path.basename(input_docx).replace(".docx", ".pdf"))
    return pdf_file

@app.route('/generate', methods=['POST'])
def generate():
    try:
        # Get form data
        student_name = request.form['student_name']
        class_name = request.form['class']
        registration_no = request.form['registration_no']
        roll_no = request.form['roll_no']
        start_year = request.form['start_year']
        end_year = request.form['end_year']
        subjects = json.loads(request.form.get('subjects', '[]'))

        template_docx = "template.docx"
        merged_doc = Document()

        # Merge subjects into one DOCX
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

            doc = DocxTemplate(template_docx)
            doc.render(context)

            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(tmp_file.name)
            tmp_file.close()

            tmp_doc = Document(tmp_file.name)
            for element in tmp_doc.element.body:
                merged_doc.element.body.append(element)

            if idx < len(subjects) - 1:
                merged_doc.add_page_break()

            os.unlink(tmp_file.name)

        # Save merged DOCX
        output_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        merged_doc.save(output_docx.name)
        output_docx.close()

        # Convert to PDF
        output_pdf = convert_docx_to_pdf(output_docx.name)
        os.unlink(output_docx.name)

        # Send PDF to frontend
        return send_file(output_pdf, as_attachment=True, download_name=f"{student_name}_frontpage.pdf")
    
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))


