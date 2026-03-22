from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
import shutil
import os
import uuid
import json
from typing import Optional

from app.parser import parse_docx
from app.formatter import generate_doc

app = FastAPI()

os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)


@app.get("/app", response_class=HTMLResponse)
def web_ui():
    with open("app/index.html") as f:
        return f.read()


@app.get("/")
def root():
    return {"message": "IEEE Formatter API is running"}


@app.post("/upload/")
async def upload(
    file: UploadFile = File(...),
    authors: Optional[str] = Form(default="[]")
):
    # Unique ID per request
    file_id = str(uuid.uuid4())

    input_path = f"uploads/{file_id}_{file.filename}"
    output_path = f"outputs/formatted_{file_id}.docx"

    # Save uploaded file
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Parse authors JSON from form field
    try:
        authors_list = json.loads(authors) if authors else []
    except (json.JSONDecodeError, TypeError):
        authors_list = []

    # Parse document and generate IEEE-formatted output
    data = parse_docx(input_path)
    data["authors"] = authors_list
    generate_doc(data, output_path)

    return {
        "message": "File processed successfully",
        "download_url": f"/download/{file_id}"
    }


@app.get("/download/{file_id}")
def download(file_id: str):
    file_path = f"outputs/formatted_{file_id}.docx"

    if not os.path.exists(file_path):
        return {"error": "File not found"}

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="IEEE_Paper.docx"
    )
