from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil, os

# Import functions from convert.py
from convert import (
    pdf_to_excel, pdf_to_word, pdf_to_ppt,
    compress_pdf, protect_pdf, unprotect_pdf, ocr_pdf, detect_mode
)

# ------------------------------
# App + CORS
# ------------------------------
app = FastAPI()

# Enable CORS (so frontend on Netlify can talk to backend on Render)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For security, replace "*" with ["https://pdftools24.netlify.app"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------------------
# API Route
# ------------------------------
@app.post("/convert")
async def convert_pdf(
    file: UploadFile,
    output: str = Form("auto"),
    password: str = Form(None)
):
    # Save uploaded PDF temporarily
    pdf_path = f"temp_{file.filename}"
    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    out_file = None
    if output == "auto":
        output = detect_mode(pdf_path)

    # Run the right conversion
    if output == "excel":
        out_file = "output.xlsx"
        pdf_to_excel(pdf_path, out_file)
    elif output == "word":
        out_file = "output.docx"
        pdf_to_word(pdf_path, out_file)
    elif output == "ppt":
        out_file = "output.pptx"
        pdf_to_ppt(pdf_path, out_file)
    elif output == "compress":
        out_file = "compressed.pdf"
        compress_pdf(pdf_path, out_file)
    elif output == "protect":
        if not password:
            os.remove(pdf_path)
            return {"error": "Password required"}
        out_file = "protected.pdf"
        protect_pdf(pdf_path, out_file, password)
    elif output == "unprotect":
        if not password:
            os.remove(pdf_path)
            return {"error": "Password required"}
        out_file = "unprotected.pdf"
        unprotect_pdf(pdf_path, out_file, password)
    elif output == "ocr":
        out_file = "ocr_output.docx"
        ocr_pdf(pdf_path, out_file)
    else:
        os.remove(pdf_path)
        return {"error": "Invalid output type"}

    # Clean up temp file
    os.remove(pdf_path)

    # Return the converted file for download
    return FileResponse(out_file, filename=out_file)
