import subprocess
from pdf2docx import Converter
import os

def rebuild_pdf(input_file, output_file):
    print(f"--- Step 1: Rebuilding {input_file} with Ghostscript ---")
    args = [
        "gs", 
        "-dNOPAUSE", "-dBATCH", "-dSAFER",
        "-sDEVICE=pdfwrite",
        f"-sOutputFile={output_file}",
        input_file
    ]
    try:
        subprocess.run(args, check=True)
        print("PDF rebuilt successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error during Ghostscript execution: {e}")
        return False
    return True

def convert_to_docx(pdf_file, docx_file):
    print(f"--- Step 2: Converting {pdf_file} to Word ---")
    try:
        cv = Converter(pdf_file)
        cv.convert(docx_file)
        # cv.convert(docx_file, multi_processing=True)
        cv.close()
        print(f"Successfully created {docx_file}")
    except Exception as e:
        print(f"Error during conversion: {e}")

if __name__ == "__main__":
    input_pdf = "book.pdf"
    rebuilt_pdf = "book_rebuilt.pdf"
    output_docx = "book_final.docx"

    if os.path.exists(input_pdf):
        if rebuild_pdf(input_pdf, rebuilt_pdf):
            convert_to_docx(rebuilt_pdf, output_docx)
    else:
        print(f"Error: {input_pdf} not found in the directory.")