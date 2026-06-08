from fastapi.responses import FileResponse
import os
from datetime import datetime
import pandas as pd

@app.post("/generate-export")
async def generate_export(data: dict):

    filename = f"export_{int(datetime.now().timestamp())}.xlsx"
    file_path = f"exports/{filename}"

    # make sure folder exists
    os.makedirs("exports", exist_ok=True)

    # create excel file
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        for sheet_name, records in data.items():
            pd.DataFrame(records).to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False
            )

    return {
        "success": True,
        "file_url": f"https://web-production-6d634.up.railway.app/files/{filename}"
    }