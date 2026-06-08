from fastapi import FastAPI
import pandas as pd
from datetime import datetime

app = FastAPI()

@app.get("/")
def root():
    return {"status": "running"}

@app.post("/generate-export")
async def generate_export(data: dict):

    filename = f"export_{int(datetime.now().timestamp())}.xlsx"

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for sheet_name, records in data.items():
            pd.DataFrame(records).to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False
            )

    return {
  "success": true,
  "file_url": "https://.../download/export_1780893151.xlsx"
}
    }
