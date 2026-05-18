from pathlib import Path
import tempfile
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from analysis_engine import analyse_file, analyse_text

app = FastAPI(title="VinQu Backend")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.get("/")
def root():
    return {"status":"ok","message":"VinQu backend is running"}

@app.get("/health")
def health():
    return {"status":"healthy"}

@app.post("/analyse/text")
async def analyse_text_endpoint(text: str = Form(...), file_name: str = Form("")):
    return analyse_text(text, file_name)

@app.post("/analyse/file")
async def analyse_file_endpoint(file: UploadFile = File(...)):
    suffix = Path(file.filename or "").suffix
    if not suffix:
        raise HTTPException(status_code=400, detail="Uploaded file has no file extension.")
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        temp_path = tmp.name
    try:
        return analyse_file(temp_path)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))
