from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from pdf2docx import Converter
from docx import Document
from openpyxl import Workbook, load_workbook
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


@app.post("/convert/docx-to-xlsx")
async def convert_docx_to_xlsx(file: UploadFile = File(...)):
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="File must be DOCX")

    input_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    input_temp.write(await file.read())
    input_temp.close()

    output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    output_temp.close()

    try:
        doc = Document(input_temp.name)

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        table_count = 0
        for table in doc.tables:
            table_count += 1
            for row in table.rows:
                ws.append([cell.text for cell in row.cells])
            ws.append([]) 

        if table_count == 0:
            raise HTTPException(status_code=400, detail="No tables found in DOCX")
        wb.save(output_temp.name)
        return FileResponse(
            output_temp.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=file.filename.replace(".docx", ".xlsx")
        )

    finally:
        os.unlink(input_temp.name)
        
@app.post("/convert/xlsx-to-docx")
async def convert_xlsx_to_docx(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="File must be XLSX or XLS")

    input_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    input_temp.write(await file.read())
    input_temp.close()

    output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    output_temp.close()

    try:
        wb = load_workbook(input_temp.name, data_only=True)
        doc = Document()

        for sheet in wb.worksheets:
            doc.add_heading(sheet.title, level=1)
            for row in sheet.iter_rows(values_only=True):
                line = "\t".join([str(cell) if cell is not None else "" for cell in row])
                doc.add_paragraph(line)
            doc.add_page_break() 

        doc.save(output_temp.name)

        return FileResponse(
            output_temp.name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.rsplit(".", 1)[0] + ".docx"
        )

    finally:
        os.unlink(input_temp.name)