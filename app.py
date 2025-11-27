from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docxtpl import DocxTemplate
from docx import Document
import tempfile
import json
import os
import threading
import time
import aspose.words as aw   # PURE PYTHON PDF CONVERTER

# -------------------------
# CONFIG
# -------------------------
# NOTE: In a real Railway environment, you should generally configure CORS to match
# your environment variables or use a more secure pattern than hardcoding.
FRONTEND_ORIGIN = os.environ.get(
    "FRONTEND_ORIGIN",
    "https://academic-r7bgbxe7k-dipros-projects-b7e275bc.vercel.app"
)

TEMPLATE_FILENAME = "template.docx"
CLEANUP_DELAY = 20

app = Flask(__name__)
# Applying CORS with the specified origin and credentials support
CORS(app, supports_credentials=True, origins=[FRONTEND_ORIGIN])


# -------------------------
# CLEANUP
# -------------------------
def remove_later(paths, delay):
    """
    Schedules file removal in a separate thread after a delay.
    This helps clean up temporary files without blocking the response.
    """
    def _worker(files):
        time.sleep(delay)
        for f in files:
            try:
                if os.path.exists(f):
                    os.remove(f)
                    print(f"Cleaned up temporary file: {f}")
            except Exception as cleanup_e:
                # Log cleanup failures but don't stop the main process
                print(f"Failed to clean up file {f}: {cleanup_e}")

    t = threading.Thread(target=_worker, args=(paths,), daemon=True)
    t.start()


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
    Generates a merged DOCX from a template, converts it to PDF using Aspose.Words,
    and returns the PDF file.
    """
    temp_files_to_clean = []
    
    try:
        # --- 1. Get and Validate Input Data ---
        # Get data from form payload
        student_name = request.form.get("student_name", "Student")
        class_name = request.form.get("class", "Unknown Class")
        registration_no = request.form.get("registration_no", "N/A")
        roll_no = request.form.get("roll_no", "N/A")
        start_year = request.form.get("start_year", "YYYY")
        end_year = request.form.get("end_year", "YYYY")

        # Parse subjects (must be a JSON array)
        subjects = []
        try:
            subjects_json = request.form.get("subjects", "[]")
            subjects = json.loads(subjects_json)
            if not isinstance(subjects, list) or not subjects:
                return jsonify({"error": "No valid subjects list provided in 'subjects' field."}), 400
        except json.JSONDecodeError:
            return jsonify({"error": "Invalid subjects JSON format."}), 400

        # Template check (assuming TEMPLATE_FILENAME exists next to app.py)
        template_path = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)
        if not os.path.exists(template_path):
            return jsonify({"error": f"template.docx missing at path: {template_path}"}), 500

        # --- 2. Process and Merge Documents ---
        merged_doc = Document()
        
        for index, subject in enumerate(subjects):
            # Context for DocxTemplate
            ctx = {
                "student_name": student_name,
                "subject": subject,
                "class": class_name,
                "registration_no": registration_no,
                "roll_no": roll_no,
                "start_year": start_year,
                "end_year": end_year,
            }

            # Render the template
            tpl = DocxTemplate(template_path)
            tpl.render(ctx)

            # Save the rendered page to a temporary DOCX file
            tmp_page = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            tpl.save(tmp_page.name)
            tmp_page.close()
            temp_files_to_clean.append(tmp_page.name)

            # Append the rendered page's content to the merged document
            page_doc = Document(tmp_page.name)
            # Use page_doc.element.body to get all content elements
            for element in page_doc.element.body:
                merged_doc.element.body.append(element)

            # Add a page break between subjects, but not after the last one
            if index < len(subjects) - 1:
                merged_doc.add_page_break()

        # Save the final merged DOCX
        merged_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        merged_doc.save(merged_file.name)
        merged_file.close()
        merged_file_path = merged_file.name
        temp_files_to_clean.append(merged_file_path)
        
        # --- 3. Convert DOCX to PDF ---
        # Aspose requires the target file to exist (or just to use its name)
        pdf_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_file.close()
        pdf_file_path = pdf_file.name
        temp_files_to_clean.append(pdf_file_path)

        # Use Aspose.Words to load the merged DOCX and save it as a PDF
        doc = aw.Document(merged_file_path)
        # Specify SaveFormat explicitly
        doc.save(pdf_file_path, aw.SaveFormat.PDF)

        # --- 4. Send File and Schedule Cleanup ---
        
        # Schedule all temporary files for delayed deletion
        remove_later(temp_files_to_clean, CLEANUP_DELAY)

        # Return the generated PDF file
        return send_file(
            pdf_file_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"{student_name}_academic_frontpages.pdf"
        )

    except Exception as e:
        # Schedule cleanup for any files created before the error occurred
        if temp_files_to_clean:
            remove_later(temp_files_to_clean, 0) # Clean immediately on failure

        # Return a server error with details for debugging
        print(f"FATAL ERROR in /generate: {e}")
        return jsonify({"error": "Server error during document processing.", "details": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    # Note: Flask's built-in server is only for development. Gunicorn is used in production.
    app.run(host="0.0.0.0", port=port)



