from fastapi import UploadFile, HTTPException
from fastapi.responses import FileResponse
import tempfile
import camelot
import subprocess
import os
import shutil

def get_soffice_binary() -> str:
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise HTTPException(status_code=503, detail="LibreOffice is not installed on the server")
    return binary

def save_upload_to_temp(file: UploadFile, suffix: str) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(file.file.read())
    tmp.close()
    return tmp.name

def response_file(path: str, filename: str, mime: str):
    return FileResponse(
        path,
        media_type=mime,
        filename=filename
    )
    
def doc_to_docx(doc_path: str) -> str:
    out_dir = tempfile.mkdtemp()
    return libreoffice_convert(doc_path, out_dir, "docx")

def extract_pdf_tables(pdf_path: str):
    tables = camelot.read_pdf(pdf_path, pages="all")
    return tables 

import subprocess
import os

def libreoffice_convert(input_path: str, output_dir: str, fmt: str) -> str:
    binary = get_soffice_binary()
    try:
        subprocess.run(
            [
                binary,
                "--headless",
                "--convert-to", fmt,
                "--outdir", output_dir,
                input_path,
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except subprocess.CalledProcessError as e:
        raise HTTPException(
            status_code=500,
            detail=f"LibreOffice conversion failed: {e.stderr.decode(errors='ignore')}",
        )
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(output_dir, f"{base}.{fmt}")
    
def libreoffice_pdf_to_docx(input_pdf: str, output_dir: str) -> str:
    return libreoffice_convert(input_pdf, output_dir, "docx")