from fastapi import FastAPI, UploadFile, File
import shutil
import os
from app.parser import parse_docx
from app.formatter import generate_doc

app = FastAPI()

# Ensure folders exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)

@app.get("/")
def root():
    return {"message": "IEEE Formatter API is running"}

@app.post("/upload/")
async def upload(file: UploadFile = File(...)):
    input_path = f"uploads/{file.filename}"
    output_path = f"outputs/formatted_{file.filename}"

    # Save uploaded file
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Process file
    data = parse_docx(input_path)
    generate_doc(data, output_path)

    return {
        "message": "File processed successfully",
        "output_file": output_path
    }