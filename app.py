from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docxtpl import DocxTemplate
from docx import Document
# Using docx2pdf for dependency clarity, but implementing the conversion via subprocess
# which is generally more robust on Linux containers.
from docx2pdf import convert 
import tempfile
import json
import os
import subprocess
import os.path
import atexit

# -------------------------
# CONFIG
# -------------------------
FRONTEND_ORIGIN = os.environ.get(
    "FRONTEND_ORIGIN",
    "https://academic-r7bgbxe7k-dipros-projects-b7e275bc.vercel.app"
)

TEMPLATE_FILENAME = "template.docx"
app = Flask(__name__)
# Applying CORS with the specified origin and credentials support
CORS(app, supports_credentials=True, origins=[FRONTEND_ORIGIN])


# -------------------------
# ROUTES
# -------------------------
@app.route("/health")
def health():
    """Simple health check endpoint."""
    return jsonify({"status": "ok"})


@app.route("/generate", methods=["POST"])
def generate():
    """
    Generates a merged DOCX from a template, converts it to PDF using 
    LibreOffice via subprocess, and returns the PDF file.
    """
    temp_files_to_clean = []
    
    # Use a temporary directory for safe cleanup of all generated files
    temp_dir = tempfile.mkdtemp()
    
    try:
        # --- 1. Get and Validate Input Data ---
        student_name = request.form.get("student_name", "Student")
        class_name = request.form.get("class", "Unknown Class")
        registration_no = request.form.get("registration_no", "N/A")
        roll_no = request.form.get("roll_no", "N/A")
        start_year = request.form.get("start_year", "YYYY")
        end_year = request.form.get("end_year", "YYYY")

        subjects = []
        try:
            subjects_json = request.form.get("subjects", "[]")
            subjects = json.loads(subjects_json)
            if not isinstance(subjects, list) or not subjects:
                return jsonify({"error": "No valid subjects list provided in 'subjects' field."}), 400
        except json.JSONDecodeError:
            return jsonify({"error": "Invalid subjects JSON format."}), 400

        template_path = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)
        if not os.path.exists(template_path):
            return jsonify({"error": f"template.docx missing at path: {template_path}"}), 500

        # --- 2. Process and Merge Documents ---
        merged_doc = Document()
        
        for index, subject in enumerate(subjects):
            ctx = {
                "student_name": student_name,
                "subject": subject,
                "class": class_name,
                "registration_no": registration_no,
                "roll_no": roll_no,
                "start_year": start_year,
                "end_year": end_year,
            }

            tpl = DocxTemplate(template_path)
            tpl.render(ctx)

            # Save rendered page to a temporary DOCX file
            # Note: We save to the temp_dir to ensure controlled cleanup
            tmp_page_path = os.path.join(temp_dir, f"page_{index}.docx")
            tpl.save(tmp_page_path)

            # Append content to merged_doc
            page_doc = Document(tmp_page_path)
            for element in page_doc.element.body:
                merged_doc.element.body.append(element)

            if index < len(subjects) - 1:
                merged_doc.add_page_break()

        # Save the final merged DOCX
        merged_file_path = os.path.join(temp_dir, "merged.docx")
        merged_doc.save(merged_file_path)
        
        # --- 3. Convert DOCX to PDF using LibreOffice ---
        # The output directory is the temporary directory
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", temp_dir,
                merged_file_path
            ], 
            check=True, # Raise an error if the command fails
            timeout=15 # Prevent hang
        )

        # LibreOffice names the output file based on the input file
        pdf_filename = "merged.pdf"
        pdf_file_path = os.path.join(temp_dir, pdf_filename)
        
        if not os.path.exists(pdf_file_path):
            raise FileNotFoundError("PDF conversion failed: output file not found.")

        # --- 4. Send File ---
        # We read the file contents and use send_file, which handles file closing
        response = send_file(
            pdf_file_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"{student_name}_academic_frontpages.pdf"
        )
        return response

    except subprocess.CalledProcessError as e:
        print(f"LibreOffice conversion failed: {e.stderr.decode()}")
        return jsonify({"error": "Document conversion failed. Check template formatting."}), 500
    
    except Exception as e:
        # Catch-all for other errors (JSON parse, file missing, etc.)
        print(f"FATAL ERROR in /generate: {e}")
        return jsonify({"error": "Server error during document processing.", "details": str(e)}), 500

    finally:
        # Clean up the entire temporary directory and all its contents
        try:
            import shutil
            shutil.rmtree(temp_dir)
            print(f"Cleaned up temporary directory: {temp_dir}")
        except Exception as cleanup_e:
            print(f"Failed to clean up temp directory {temp_dir}: {cleanup_e}")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    # Note: Use Gunicorn in production (via Dockerfile CMD). This is for local testing.
    app.run(host="0.0.0.0", port=port)


