import os
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from datetime import datetime

app = FastAPI(title="108TS Premium Web App")

# Enable CORS (Good for development)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Set up paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Ensure required directories exist
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# API Endpoints
@app.get("/api/config")
async def get_config():
    """Return default configuration analogous to 108TS config"""
    return {
        "index": "NIFTY",
        "data_date": datetime.today().strftime("%Y-%m-%d"),
        "mid_strike": "24000",
        "auto_mid": True,
        "rows": 3,
        "gap": 50,
        "target_sh": "Sheet1",
        "dynamic_name": True,
        "template": "108TSA"
    }

@app.post("/api/sync")
async def sync_data(
    index: str = Form(...),
    expiry: str = Form(...),
    data_date: str = Form(...),
    target_sh: str = Form(...),
    dynamic_name: bool = Form(True),
    template_name: str = Form("108TSA"),
    mid_strike: float = Form(24000),
    auto_mid: bool = Form(True),
    rows: int = Form(3),
    gap: int = Form(50)
):
    """
    Endpoint that handles Step 1 & Step 2 for the Web UI.
    1. Fetches Bhavcopy
    2. Modifies Template via openpyxl
    3. Saves to Output directory
    4. Returns Download link
    """
    try:
        from data_engine import DataEngine
        from excel_engine import ExcelEngine
        
        d_engine = DataEngine()
        e_engine = ExcelEngine()
        
        # Calculate dynamic sheet name
        if dynamic_name:
            try:
                d = datetime.strptime(data_date, "%Y-%m-%d")
                e = datetime.strptime(expiry, "%Y-%m-%d")
                target_sh = f"{d.strftime('%d %b')} Exp {e.strftime('%d %b')}"
            except: pass
            
        cfg = {
            "index": index,
            "expiry": expiry,
            "data_date": data_date,
            "target_sh": target_sh,
            "template": template_name,
            "tpl_file": os.path.join(BASE_DIR, "Nifty50.xlsx"),
            "rows": rows,
            "gap": gap
        }
        
        if auto_mid:
            atm = d_engine.fetch_atm_strike(index, data_date)
            cfg['mid_strike'] = float(atm) if atm else mid_strike
        else:
            cfg['mid_strike'] = float(mid_strike)
            
        data = d_engine.fetch_bhavcopy(data_date)
        if not data:
            raise HTTPException(status_code=400, detail="Failed to fetch Bhavcopy data.")
            
        out_filename = e_engine.run_sync(cfg, data, OUTPUT_DIR)
        filename = out_filename
        return {"status": "success", "message": "Data synchronized successfully", "download_url": f"/api/download/{filename}"}
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(path=file_path, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    raise HTTPException(status_code=404, detail="File not found")

# Serve Frontend static files
try:
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
except Exception as e:
    print(f"Warning: Ensure {STATIC_DIR} has an index.html file.")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8008)
