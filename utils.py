from fastapi import UploadFile, HTTPException
from fastapi.responses import FileResponse
import tempfile
import camelot
import subprocess
import os

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
    libreoffice_convert(doc_path, out_dir, "docx")
    return os.path.join(out_dir, os.path.basename(doc_path).replace(".doc", ".docx"))

def extract_pdf_tables(pdf_path: str):
    tables = camelot.read_pdf(pdf_path, pages="all")
    return tables 

import subprocess
import os

def libreoffice_convert(input_path: str, output_dir: str, fmt: str) -> str:
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--infilter=writer_pdf_import",
            "--convert-to", fmt,
            "--outdir", output_dir,
            input_path,
        ],
        check=True
    )

    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(output_dir, f"{base}.{fmt}")
    
def libreoffice_pdf_to_docx(input_pdf: str, output_dir: str) -> str:
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--infilter=writer_pdf_import",
            "--convert-to", "docx",
            "--outdir", output_dir,
            input_pdf
        ],
        check=True
    )

    base = os.path.splitext(os.path.basename(input_pdf))[0]
    return os.path.join(output_dir, base + ".docx")
