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
BUBBLE_BASE_URL = os.getenv("BUBBLE_BASE_URL")  # MUST end in /obj
BUBBLE_API_TOKEN = os.getenv("BUBBLE_API_TOKEN")

print("BASE URL:", BUBBLE_BASE_URL)
print("TOKEN EXISTS:", BUBBLE_API_TOKEN is not None)

# IMPORTANT: Bubble Data API types (must match your Bubble exactly)
DATA_TYPES = [
    "user",
    "servicerequest",
    "driverapplication"
]

# -----------------------------
# FETCH ALL RECORDS FROM BUBBLE
# -----------------------------
def fetch_all_records(data_type):
    headers = {
        "Authorization": f"Bearer {BUBBLE_API_TOKEN}",
        "Content-Type": "application/json"
    }

    results = []
    cursor = 0

    while True:
        response = requests.get(
            f"{BUBBLE_BASE_URL}/{data_type}",
            headers=headers,
            params={"cursor": cursor}
        )

        if response.status_code != 200:
            raise Exception(f"{data_type} failed: {response.text}")

        payload = response.json().get("response", {})
        records = payload.get("results", [])

        results.extend(records)

        remaining = payload.get("remaining", 0)

        if remaining == 0:
            break

        # safer pagination
        cursor = payload.get("cursor", cursor + len(records))

    return results

# -----------------------------
# DEBUG ENDPOINT
# -----------------------------
@app.get("/debug-env")
def debug_env():
    return {
        "base_url": BUBBLE_BASE_URL,
        "token_exists": BUBBLE_API_TOKEN is not None
    }

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

                if records and len(records) > 0:
                    df = pd.DataFrame(records)
                else:
                    df = pd.DataFrame([{"empty": True}])

                df.to_excel(writer, sheet_name=data_type[:31], index=False)

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
        "file_url": f"/files/{filename}"
    }

# -----------------------------
# FILE DOWNLOAD ENDPOINT
# -----------------------------
@app.get("/files/{filename}")
def get_file(filename: str):

    file_path = f"exports/{filename}"

    if not os.path.exists(file_path):
        return {"error": "file not found"}

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# TEST BUBBLE CONNECTION
# -----------------------------
@app.get("/test-bubble")
def test_bubble():

    headers = {
        "Authorization": f"Bearer {BUBBLE_API_TOKEN}"
    }

    response = requests.get(
        f"{BUBBLE_BASE_URL}/servicerequest",
        headers=headers
    )

    return {
        "status_code": response.status_code,
        "text": response.text[:500]
    }