from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import tempfile
import os
import shutil
from generate_mvr import read_excel, fill_template

app = FastAPI()

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, "approved_file.docx")

@app.post("/generate_mvr")
async def generate_mvr(excel_file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        shutil.copyfileobj(excel_file.file, f)
        excel_path = f.name

    try:
        data = read_excel(excel_path)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as out:
            output_path = out.name

        fill_template(data, TEMPLATE, output_path)

        return FileResponse(
            output_path,
            filename="Method_Validation_Report.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    finally:
        os.unlink(excel_path)

@app.get("/")
async def home():
    return {"message": "MVR Generator is live. POST your Excel to /generate_mvr"}