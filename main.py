from fastapi import FastAPI
from api.document_converter import router as document_converter_router

app = FastAPI(title="Document Converter API")

app.include_router(document_converter_router)

@app.get("/")
def root():
    return {"status": "ok"}
