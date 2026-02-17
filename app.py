from flask import Flask, request, jsonify
import os
import pandas as pd
from docx import Document

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ---------------------------
# Helper: Process Excel Files
# ---------------------------
def analyze_excel(path):
    df = pd.read_excel(path)

    summary = {
        "rows": len(df),
        "columns": list(df.columns),
        "numeric_summary": df.describe().to_dict()
    }
    return summary


# ---------------------------
# Helper: Process Word Files
# ---------------------------
def analyze_word(path):
    doc = Document(path)
    text = "\n".join([p.text for p in doc.paragraphs])

    word_count = len(text.split())
    char_count = len(text)

    return {
        "word_count": word_count,
        "char_count": char_count,
        "preview": text[:200]  # first 200 chars
    }


# ---------------------------
# Upload Endpoint
# ---------------------------
@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    filename = file.filename
    path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(path)

    # Determine file type
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        result = analyze_excel(path)
    elif filename.endswith(".docx"):
        result = analyze_word(path)
    else:
        return jsonify({"error": "Unsupported file type"}), 400

    return jsonify({
        "filename": filename,
        "analysis": result
    })


@app.route("/", methods=["GET"])
def home():
    return "SaaS File Analytics API is running!"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
