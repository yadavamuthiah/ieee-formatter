# ============================================================
# app.py  -  the web server for the IEEE Paper Formatter
# ============================================================
# What this file does:
#   1. Shows the upload page (templates/index.html) at  http://localhost:5000
#   2. Receives the uploaded .docx file from that page
#   3. Sends it through format_ieee() (in ieee_formatter.py)
#   4. Sends the formatted .docx back so the browser downloads it
#
# How to run:   py -3.13 app.py     (then open localhost:5000)
# ============================================================

import io  # lets us keep files in memory instead of saving them on disk

# Flask is the small web framework that runs the website
from flask import Flask, render_template, request, send_file, jsonify

# format_ieee is our formatting engine (see ieee_formatter.py)
from ieee_formatter import format_ieee

# Create the web application
app = Flask(__name__)

# Reject uploads bigger than 25 MB so the server is not overloaded
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


# ---------- Route 1: the home page ----------
# When someone opens http://localhost:5000 this function runs
# and returns the HTML page from the "templates" folder.
@app.get("/")
def index():
    return render_template("index.html")


# ---------- Route 2: the formatting endpoint ----------
# The web page sends the uploaded file here (POST /format).
@app.post("/format")
def format_document():
    # Get the uploaded file from the form (the field is named "file")
    f = request.files.get("file")

    # Check a file was sent and that it ends with .docx
    if not f or not f.filename.lower().endswith(".docx"):
        # 400 = "bad request"; the page shows this message to the user
        return jsonify(error="Please upload a .docx file."), 400

    try:
        # "out" is an empty in-memory file that will hold the result
        out = io.BytesIO()

        # Read the upload into memory and format it into "out"
        # verbose=False stops it printing the report in the terminal
        format_ieee(io.BytesIO(f.read()), out, verbose=False)
    except Exception as e:
        # If anything fails (for example a damaged file), tell the user
        return jsonify(error=f"Could not process this file: {e}"), 500

    # Go back to the start of the in-memory file so it can be sent
    out.seek(0)

    # Name the download:  MyPaper.docx  ->  MyPaper_IEEE.docx
    name = f.filename.rsplit(".", 1)[0] + "_IEEE.docx"

    # Send the file to the browser as a download
    return send_file(
        out,
        as_attachment=True,
        download_name=name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------- Start the server ----------
# This block only runs when you start the file directly (py -3.13 app.py)
if __name__ == "__main__":
    print("Open http://localhost:5000 in your browser")
    app.run(host="127.0.0.1", port=5000, debug=False)
