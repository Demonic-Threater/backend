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
FRONTEND_ORIGIN = os.environ.get(
    "FRONTEND_ORIGIN",
    "https://academic-r7bgbxe7k-dipros-projects-b7e275bc.vercel.app"
)

TEMPLATE_FILENAME = "template.docx"
CLEANUP_DELAY = 20

app = Flask(__name__)
CORS(app, supports_credentials=True, origins=[FRONTEND_ORIGIN])


# -------------------------
# CLEANUP
# -------------------------
def remove_later(paths, delay):
    def _worker(files):
        time.sleep(delay)
        for f in files:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except:
                pass

    t = threading.Thread(target=_worker, args=(paths,), daemon=True)
    t.start()


# -------------------------
# ROUTES
# -------------------------
@app.route("/generate", methods=["POST"])
def generate():
    try:
        # -----------------------------
        # Parse form data safely
        # -----------------------------
        student_name = request.form.get("student_name", "Student")
        class_name = request.form.get("class", "")
        registration_no = request.form.get("registration_no", "")
        roll_no = request.form.get("roll_no", "")
        start_year = request.form.get("start_year", "")
        end_year = request.form.get("end_year", "")

        # Subjects parsing
        try:
            subjects = json.loads(request.form.get("subjects", "[]"))
            if not isinstance(subjects, list) or not subjects:
                return jsonify({"error": "Invalid or empty subjects list"}), 400
        except Exception as e:
            return jsonify({"error": "Invalid subjects JSON", "details": str(e)}), 400

        # -----------------------------
        # Check template
        # -----------------------------
        template_path = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)
        if not os.path.exists(template_path):
            return jsonify({"error": "template.docx missing"}), 500

        # -----------------------------
        # Merge DOCX
        # -----------------------------
        merged_doc = Document()
        temp_files = []

        for idx, subject in enumerate(subjects):
            try:
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

                tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                tpl.save(tmp_file.name)
                tmp_file.close()
                temp_files.append(tmp_file.name)

                page_doc = Document(tmp_file.name)
                for element in page_doc.element.body:
                    merged_doc.element.body.append(element)

                if idx < len(subjects) - 1:
                    merged_doc.add_page_break()

            except Exception as e:
                logging.exception(f"Error processing subject '{subject}'")
                return jsonify({"error": f"Failed to process subject '{subject}'", "details": str(e)}), 500

        merged_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        merged_doc.save(merged_file.name)
        merged_file.close()
        temp_files.append(merged_file.name)

        # -----------------------------
        # Convert DOCX → PDF safely
        # -----------------------------
        pdf_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_file.close()
        try:
            doc = aw.Document(merged_file.name)
            doc.save(pdf_file.name)
            temp_files.append(pdf_file.name)
        except Exception as e:
            logging.exception("PDF conversion failed")
            remove_later(temp_files, 0)
            return jsonify({"error": "PDF conversion failed", "details": str(e)}), 500

        # -----------------------------
        # Cleanup old temp files asynchronously
        # -----------------------------
        remove_later(temp_files[:-1], CLEANUP_DELAY)  # keep PDF for download

        return send_file(
            pdf_file.name,
            as_attachment=True,
            download_name=f"{student_name}_frontpage.pdf"
        )

    except Exception as e:
        logging.exception("Unexpected server error in /generate")
        return jsonify({"error": "Server error", "details": str(e)}), 500



