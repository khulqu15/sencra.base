from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from pdf2docx import Converter
import tempfile
import os

app = FastAPI()

@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File must be PDF")

    input_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    input_temp.write(await file.read())
    input_temp.close()
    output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    output_temp.close()

    try:
        cv = Converter(input_temp.name)
        cv.convert(output_temp.name, start=0, end=None)
        cv.close()
        return FileResponse(output_temp.name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=file.filename.replace(".pdf", ".docx"))

    finally:
        os.unlink(input_temp.name)
