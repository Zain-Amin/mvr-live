from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import tempfile
import os
import shutil
from generate_mvr import read_excel, build

app = FastAPI()

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shaigan_template.docx")

@app.post("/generate_mvr")
async def generate_mvr(excel_file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        shutil.copyfileobj(excel_file.file, f)
        excel_path = f.name

    output_path = None
    try:
        data = read_excel(excel_path)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as out:
            output_path = out.name

        build(data, TEMPLATE, output_path)

        return FileResponse(
            output_path,
            filename="Method_Validation_Report.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    finally:
        if os.path.exists(excel_path):
            os.unlink(excel_path)

@app.get("/")
async def home():
    return {"message": "MVR Generator is live. POST your Excel to /generate_mvr"}