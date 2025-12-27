from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from api.document_converter import router as document_converter_router

app = FastAPI(title="SENCRA BASE")
app.include_router(document_converter_router)

class ConversionException(Exception):
    def __init__(self, message: str, detail: str | None = None, status: int = 422):
        self.message = message
        self.detail = detail
        self.status = status


@app.exception_handler(ConversionException)
async def conversion_exception_handler(request: Request, exc: ConversionException):
    return JSONResponse(
        status_code=exc.status,
        content={"error": exc.message, "detail": exc.detail},
    )

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content={"error": exc.detail},
    )

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"error": "Internal server error", "detail": str(exc)},
    )

@app.get("/")
def root():
    return {"status": "ok"}
