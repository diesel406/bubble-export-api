from fastapi import FastAPI
from fastapi.responses import FileResponse
import pandas as pd
import requests
import os
from datetime import datetime

app = FastAPI()

# -----------------------------
# BUBBLE CONFIG
# -----------------------------
BUBBLE_BASE_URL = os.getenv("BUBBLE_BASE_URL")
BUBBLE_API_TOKEN = os.getenv("BUBBLE_API_TOKEN")

DATA_TYPES = [
    "User",
    "ServiceRequest",
    "DriverApplication"
]

# -----------------------------
# FETCH ALL RECORDS FROM BUBBLE
# -----------------------------
def fetch_all_records(data_type):
    headers = {
        "Authorization": f"Bearer {BUBBLE_API_TOKEN}"
    }

    results = []
    cursor = 0

    while True:

        response = requests.get(
            f"{BUBBLE_BASE_URL}/{data_type}",
            headers=headers,
            params={"cursor": cursor}
        )

        response.raise_for_status()

        payload = response.json()["response"]

        records = payload["results"]

        results.extend(records)

        remaining = payload.get("remaining", 0)

        if remaining == 0:
            break

        cursor += len(records)

    return results

# -----------------------------
# GENERATE EXPORT
# -----------------------------
@app.post("/generate-export")
async def generate_export():

    os.makedirs("exports", exist_ok=True)

    filename = f"export_{int(datetime.now().timestamp())}.xlsx"
    file_path = f"exports/{filename}"

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:

        for data_type in DATA_TYPES:

            try:
                records = fetch_all_records(data_type)

                if records:
                    pd.DataFrame(records).to_excel(
                        writer,
                        sheet_name=data_type[:31],
                        index=False
                    )

            except Exception as e:
                pd.DataFrame([
                    {"error": str(e)}
                ]).to_excel(
                    writer,
                    sheet_name=f"{data_type}_ERROR"[:31],
                    index=False
                )

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