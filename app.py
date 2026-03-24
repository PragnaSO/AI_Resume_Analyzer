import os
from pathlib import Path

from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

from utils import analyze_resume_against_jd, extract_text_from_resume


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
ALLOWED_EXTENSIONS = {"pdf", "docx", "doc"}

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)

UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    if "resume" not in request.files:
        return render_template("index.html", error="Please upload a resume file (PDF, DOCX, or DOC)."), 400

    file = request.files["resume"]
    job_description = request.form.get("job_description", "").strip()

    if file.filename == "":
        return render_template("index.html", error="No file selected."), 400

    if not job_description:
        return render_template("index.html", error="Please enter a job description."), 400

    if not allowed_file(file.filename):
        return render_template("index.html", error="Only PDF, DOCX, and DOC files are supported."), 400

    safe_name = secure_filename(file.filename)
    save_path = os.path.join(app.config["UPLOAD_FOLDER"], safe_name)
    file.save(save_path)

    try:
        resume_text = extract_text_from_resume(save_path)
    except ValueError as error:
        return render_template("index.html", error=str(error)), 400
    except Exception:
        return render_template(
            "index.html",
            error="Unable to read the file. Please upload a valid PDF, DOCX, or DOC document.",
        ), 400

    if not resume_text.strip():
        return render_template("index.html", error="No readable text found in the uploaded file."), 400

    result = analyze_resume_against_jd(resume_text, job_description)
    return render_template("index.html", result=result, job_description=job_description)


if __name__ == "__main__":
    app.run(debug=True)
