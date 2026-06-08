from fastapi import FastAPI
from fastapi.responses import FileResponse
import pandas as pd
import os
from datetime import datetime

app = FastAPI()

# -----------------------------
# CREATE EXPORT + RETURN URL
# -----------------------------
@app.post("/generate-export")
async def generate_export(data: dict):

    # make sure folder exists
    os.makedirs("exports", exist_ok=True)

    # create unique filename
    filename = f"export_{int(datetime.now().timestamp())}.xlsx"
    file_path = f"exports/{filename}"

    # write Excel file with multiple sheets
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        for sheet_name, records in data.items():
            pd.DataFrame(records).to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False
            )

    # return Bubble-friendly response
    return {
        "success": True,
        "file_url": f"https://web-production-6d634.up.railway.app/files/{filename}"
    }


# -----------------------------
# FILE DOWNLOAD ENDPOINT
# -----------------------------
@app.get("/files/{filename}")
def get_file(filename: str):

    file_path = f"exports/{filename}"

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )