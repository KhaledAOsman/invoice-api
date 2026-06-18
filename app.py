import os
import uuid
import tempfile
from flask import Flask, request, send_file, jsonify, render_template

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER = tempfile.gettempdir()

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() == "pdf"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if "pdf" not in request.files:
        return jsonify({"error": "لم يتم إرسال ملف"}), 400

    file = request.files["pdf"]

    if not file.filename or not allowed_file(file.filename):
        return jsonify({"error": "يجب أن يكون الملف بصيغة PDF"}), 400

    uid = str(uuid.uuid4())[:8]
    pdf_path   = os.path.join(UPLOAD_FOLDER, f"{uid}_input.pdf")
    excel_path = os.path.join(UPLOAD_FOLDER, f"{uid}_output.xlsx")

    file.save(pdf_path)

    try:
        from invoice_parser import extract_all_invoices, to_excel_merged
        all_invoices = extract_all_invoices(pdf_path)

        if not all_invoices:
            return jsonify({"error": "لم يتم العثور على فواتير في الملف"}), 422

        to_excel_merged(all_invoices, excel_path)

        return send_file(
            excel_path,
            as_attachment=True,
            download_name="invoices_output.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": f"خطأ أثناء المعالجة: {str(e)}"}), 500
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
