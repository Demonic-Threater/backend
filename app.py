from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docxtpl import DocxTemplate
from docx import Document
import tempfile
import json
import os
import subprocess
import shutil
import threading
import time
import traceback

# -------------------------
# Config
# -------------------------
FRONTEND_ORIGIN = os.environ.get("FRONTEND_ORIGIN", "*")  # set to your frontend domain in production
TEMPLATE_FILENAME = os.environ.get("TEMPLATE_FILENAME", "template.docx")
CLEANUP_DELAY = int(os.environ.get("CLEANUP_DELAY", "30"))  # seconds to wait before deleting temp files

app = Flask(__name__)
CORS(app, origins=FRONTEND_ORIGIN)


# -------------------------
# Helpers
# -------------------------
def find_soffice_executable():
    """Return path to 'soffice' or 'libreoffice' if available, else None."""
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def background_remove(files, delay=CLEANUP_DELAY):
    """Remove given file paths after delay seconds in a background thread."""
    def _worker(paths, wait):
        time.sleep(wait)
        for p in paths:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
    t = threading.Thread(target=_worker, args=(files, delay), daemon=True)
    t.start()


# -------------------------
# Routes
# -------------------------
@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/generate", methods=["POST"])
def generate():
    try:
        # Required fields (provide defaults as needed)
        student_name = request.form.get("student_name", "").strip() or "Student"
        class_name = request.form.get("class", "").strip()
        registration_no = request.form.get("registration_no", "").strip()
        roll_no = request.form.get("roll_no", "").strip()
        start_year = request.form.get("start_year", "").strip()
        end_year = request.form.get("end_year", "").strip()

        # subjects expected as JSON array string
        subjects_raw = request.form.get("subjects", "[]")
        try:
            subjects = json.loads(subjects_raw)
            if not isinstance(subjects, list):
                raise ValueError("subjects must be a JSON array")
        except Exception as e:
            return jsonify({"error": "Invalid subjects JSON", "details": str(e)}), 400

        if len(subjects) == 0:
            return jsonify({"error": "No subjects provided"}), 400

        # Template path
        template_path = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)
        if not os.path.exists(template_path):
            return jsonify({"error": "Template not found on server", "template_path": template_path}), 500

        # Create merged doc
        merged_doc = Document()

        tmp_docx_files = []
        for idx, subject in enumerate(subjects):
            ctx = {
                "student_name": student_name,
                "subject": subject,
                "class": class_name,
                "registration_no": registration_no,
                "roll_no": roll_no,
                "start_year": start_year,
                "end_year": end_year
            }

            tpl = DocxTemplate(template_path)
            tpl.render(ctx)

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            tpl.save(tmp.name)
            tmp.close()
            tmp_docx_files.append(tmp.name)

            # Append content
            tmp_doc = Document(tmp.name)
            for element in tmp_doc.element.body:
                merged_doc.element.body.append(element)

            if idx < len(subjects) - 1:
                # add a page break between pages
                merged_doc.add_page_break()

        # Save merged docx to temp file
        merged_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        merged_doc.save(merged_docx.name)
        merged_docx.close()

        # Prepare output pdf path
        out_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        out_pdf.close()

        # Convert with LibreOffice / soffice
        soffice = find_soffice_executable()
        if not soffice:
            # Cleanup docx files we created
            files_to_remove = tmp_docx_files + [merged_docx.name]
            background_remove(files_to_remove, delay=5)
            return (
                jsonify(
                    {
                        "error": "LibreOffice (soffice/libreoffice) not found on server. PDF conversion unavailable."
                    }
                ),
                500,
            )

        # LibreOffice outputs pdf to same directory as input; call with outdir
        outdir = os.path.dirname(out_pdf.name)
        try:
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    outdir,
                    merged_docx.name,
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60,
            )
        except subprocess.CalledProcessError as e:
            # conversion failed
            details = e.stderr.decode("utf-8", errors="ignore") if e.stderr else str(e)
            files_to_remove = tmp_docx_files + [merged_docx.name, out_pdf.name]
            background_remove(files_to_remove, delay=5)
            app.logger.error("LibreOffice conversion failed: %s", details)
            return jsonify({"error": "PDF conversion failed", "details": details}), 500
        except subprocess.TimeoutExpired:
            files_to_remove = tmp_docx_files + [merged_docx.name, out_pdf.name]
            background_remove(files_to_remove, delay=5)
            return jsonify({"error": "PDF conversion timed out"}), 500

        # LibreOffice names pdf identical to docx basename
        expected_pdf = os.path.splitext(merged_docx.name)[0] + ".pdf"
        if not os.path.exists(expected_pdf):
            # Maybe conversion failed silently
            files_to_remove = tmp_docx_files + [merged_docx.name, out_pdf.name]
            background_remove(files_to_remove, delay=5)
            return jsonify({"error": "Converted PDF not found", "expected": expected_pdf}), 500

        # Move/rename expected_pdf to out_pdf.name
        try:
            os.replace(expected_pdf, out_pdf.name)
        except Exception:
           


