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
# HELPERS
# -------------------------
def find_soffice():
    for exe in ["soffice", "libreoffice"]:
        p = shutil.which(exe)
        if p:
            return p
    return None


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
@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/generate", methods=["POST"])
def generate():
    try:
        student_name = request.form.get("student_name", "Student")
        class_name = request.form.get("class", "")
        registration_no = request.form.get("registration_no", "")
        roll_no = request.form.get("roll_no", "")
        start_year = request.form.get("start_year", "")
        end_year = request.form.get("end_year", "")

        # Parse subjects
        try:
            subjects = json.loads(request.form.get("subjects", "[]"))
            if not isinstance(subjects, list):
                raise ValueError
        except:
            return jsonify({"error": "Invalid subjects JSON"}), 400

        if len(subjects) == 0:
            return jsonify({"error": "No subjects provided"}), 400

        # Template existence check
        template_path = os.path.join(os.path.dirname(__file__), TEMPLATE_FILENAME)
        if not os.path.exists(template_path):
            return jsonify({"error": "template.docx missing on server"}), 500

        # Build merged doc
        merged_doc = Document()
        temp_pages = []

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

            tmp_page = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            tpl.save(tmp_page.name)
            tmp_page.close()

            temp_pages.append(tmp_page.name)

            page_doc = Document(tmp_page.name)
            for element in page_doc.element.body:
                merged_doc.element.body.append(element)

            if index < len(subjects) - 1:
                merged_doc.add_page_break()

        # Save merged docx
        merged_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        merged_doc.save(merged_docx.name)
        merged_docx.close()

        # Create output PDF temp path
        pdf_out = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_out.close()

        # Convert using LibreOffice
        soffice = find_soffice()
        if not soffice:
            remove_later(temp_pages + [merged_docx.name, pdf_out.name], CLEANUP_DELAY)
            return jsonify({"error": "LibreOffice not installed"}), 500

        outdir = os.path.dirname(pdf_out.name)
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
                timeout=40,
            )
        except Exception as e:
            remove_later(temp_pages + [merged_docx.name, pdf_out.name], CLEANUP_DELAY)
            return jsonify({"error": "PDF conversion failed", "details": str(e)}), 500

        # Expected output path
        expected_pdf = os.path.splitext(merged_docx.name)[0] + ".pdf"

        if not os.path.exists(expected_pdf):
            remove_later(temp_pages + [merged_docx.name, pdf_out.name], CLEANUP_DELAY)
            return jsonify({"error": "PDF missing after convert"}), 500

        # Move generated PDF to final path
        os.replace(expected_pdf, pdf_out.name)

        # Cleanup temp DOCX after return
        remove_later(temp_pages + [merged_docx.name], CLEANUP_DELAY)

        return send_file(
            pdf_out.name,
            as_attachment=True,
            download_name=f"{student_name}_frontpage.pdf"
        )

    except Exception as e:
        return jsonify({"error": "Server error", "details": str(e)}), 500


# -------------------------
# MAIN
# -------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)






