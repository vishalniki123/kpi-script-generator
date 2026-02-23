from flask import Flask, render_template, request, jsonify, send_file
import os
import zipfile
from kpi_generator import run_generator

app = Flask(__name__)
app.secret_key = "kpi-secret"

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
PSQ_OUTPUT = os.path.join(OUTPUT_DIR, "psq")
SQD_OUTPUT = os.path.join(OUTPUT_DIR, "sqd")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PSQ_OUTPUT, exist_ok=True)
os.makedirs(SQD_OUTPUT, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = {}
        for f in request.files:
            path = os.path.join(UPLOAD_DIR, request.files[f].filename)
            request.files[f].save(path)
            files[f] = path

        run_generator(
            files["kpi_excel"],
            files["psq1"],
            files["psq2"],
            files["sqd1"],
            files["sqd2"],
            PSQ_OUTPUT,
            SQD_OUTPUT
        )

        zip_path = os.path.join(OUTPUT_DIR, "KPI_Output.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            for folder in [PSQ_OUTPUT, SQD_OUTPUT]:
                for file in os.listdir(folder):
                    z.write(os.path.join(folder, file), arcname=file)

        return jsonify({
            "status": "success",
            "files": {
                "psq": os.listdir(PSQ_OUTPUT),
                "sqd": os.listdir(SQD_OUTPUT)
            }
        })

    return render_template("index.html")

@app.route("/download")
def download():
    zip_path = os.path.join(OUTPUT_DIR, "KPI_Output.zip")
    return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)from flask import Flask, render_template, request, jsonify, send_file
import os
import zipfile
from kpi_generator import run_generator

app = Flask(__name__)
app.secret_key = "kpi-secret"

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
PSQ_OUTPUT = os.path.join(OUTPUT_DIR, "psq")
SQD_OUTPUT = os.path.join(OUTPUT_DIR, "sqd")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PSQ_OUTPUT, exist_ok=True)
os.makedirs(SQD_OUTPUT, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = {}
        for f in request.files:
            path = os.path.join(UPLOAD_DIR, request.files[f].filename)
            request.files[f].save(path)
            files[f] = path

        run_generator(
            files["kpi_excel"],
            files["psq1"],
            files["psq2"],
            files["sqd1"],
            files["sqd2"],
            PSQ_OUTPUT,
            SQD_OUTPUT
        )

        zip_path = os.path.join(OUTPUT_DIR, "KPI_Output.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            for folder in [PSQ_OUTPUT, SQD_OUTPUT]:
                for file in os.listdir(folder):
                    z.write(os.path.join(folder, file), arcname=file)

        return jsonify({
            "status": "success",
            "files": {
                "psq": os.listdir(PSQ_OUTPUT),
                "sqd": os.listdir(SQD_OUTPUT)
            }
        })

    return render_template("index.html")

@app.route("/download")
def download():
    zip_path = os.path.join(OUTPUT_DIR, "KPI_Output.zip")
    return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
