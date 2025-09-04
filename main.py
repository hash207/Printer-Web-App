import os
import subprocess as sp
from flask import Flask, render_template, request, redirect, url_for, flash, get_flashed_messages

app = Flask(__name__)
app.secret_key = "your_secret_key_here"  # Set a secret key for flashing

DOCS_FOLDER = os.path.join(os.path.dirname(__file__), "Documints")
os.makedirs(DOCS_FOLDER, exist_ok=True)

def print_document_linux(file_path, printer_name=None, copies=1):
    """
    Prints a document to a specified or default printer on Linux.

    Args:
        file_path (str): The full path to the document to be printed.
        printer_name (str, optional): The name of the printer to use. 
                                      If None, the default printer is used.
        copies (int, optional): The number of copies to print. Defaults to 1.
    """
    command = ["lp"]

    if printer_name:
        command.extend(["-d", printer_name])

    if copies > 1:
        command.extend(["-n", str(copies)])

    command.append(file_path)

    try:
        sp.run(command, check=True, capture_output=True, text=True)
        print(f"Document '{file_path}' sent to printer successfully.")
    except sp.CalledProcessError as e:
        print(f"Error printing document: {e}")
        print(f"Stderr: {e.stderr}")
    except FileNotFoundError:
        print("Error: 'lp' command not found. Ensure CUPS is installed.")

def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Converts a DOCX file to PDF using LibreOffice.
    """
    try:
        sp.run([
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", os.path.dirname(pdf_path),
            docx_path
        ], check=True, capture_output=True)
        # LibreOffice names the PDF as <basename>.pdf in the same folder
        # Ensure the output file is named as expected
        generated_pdf = os.path.splitext(docx_path)[0] + ".pdf"
        if generated_pdf != pdf_path:
            os.rename(generated_pdf, pdf_path)
    except Exception as e:
        raise RuntimeError(f"LibreOffice conversion failed: {e}")

@app.route("/")
def home():
    messages = get_flashed_messages()
    return render_template("home.html", messages=messages)

@app.route("/print", methods=["POST"])
def print_document():
    uploaded_file = request.files.get("document")
    copies = request.form.get("copies")
    if uploaded_file:
        save_path = os.path.join(DOCS_FOLDER, uploaded_file.filename)
        uploaded_file.save(save_path)
        # Convert .docx to .pdf if needed
        if uploaded_file.filename.lower().endswith(".docx"):
            pdf_path = os.path.splitext(save_path)[0] + ".pdf"
            try:
                convert_docx_to_pdf(save_path, pdf_path)
                file_to_print = pdf_path
            except Exception as e:
                flash(f"Error converting DOCX to PDF: {e}")
                return redirect(url_for("home"))
        else:
            file_to_print = save_path
        # Call the print function
        try:
            print_document_linux(file_to_print, copies=int(copies))
            flash("Document sent to printer.")
        except Exception as e:
            flash(f"Error printing document: {e}")
            return redirect(url_for("home"))
    else:
        flash("No document uploaded.")
    return redirect(url_for("home"))

if __name__ == "__main__":
    app.run("0.0.0.0")