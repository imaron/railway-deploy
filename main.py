{\rtf1\ansi\ansicpg1252\cocoartf2822
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 from fastapi import FastAPI, File, UploadFile\
from fastapi.responses import JSONResponse\
import uvicorn\
import os\
import requests\
import tempfile\
import shutil\
from optimize_schedules_with_sanity import (\
    read_cost_pref_hours_caps,\
    solve_cpsat,\
    write_solution\
)\
\
app = FastAPI(title="Schedule Optimizer API")\
\
@app.post("/run")\
async def run_schedule_optimizer(file: UploadFile = File(...)):\
    try:\
        # Save uploaded file\
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:\
            shutil.copyfileobj(file.file, tmp_in)\
            input_path = tmp_in.name\
\
        # Prepare output file path\
        output_path = input_path.replace(".xlsx", "_Solved.xlsx")\
\
        # Run your existing script logic\
        costs, prefs, hours, lam, shift_caps, hour_caps = read_cost_pref_hours_caps(input_path)\
        sol, obj = solve_cpsat(costs, prefs, hours, lam, shift_caps, hour_caps)\
        write_solution(input_path, sol, obj, output_path, costs, hours, hour_caps)\
\
        # Upload solved file to tmpfiles.org\
        with open(output_path, "rb") as f:\
            r = requests.post("https://tmpfiles.org/api/v1/upload", files=\{"file": f\})\
            if r.status_code != 200:\
                raise RuntimeError(f"Upload failed: \{r.text\}")\
            download_url = r.json()["data"]["url"]\
\
        # Cleanup local temp files\
        os.remove(input_path)\
        os.remove(output_path)\
\
        return JSONResponse(content=\{\
            "status": "success",\
            "objective": obj,\
            "download_url": download_url\
        \})\
\
    except Exception as e:\
        return JSONResponse(status_code=500, content=\{\
            "status": "error",\
            "message": str(e)\
        \})\
\
\
if __name__ == "__main__":\
    uvicorn.run(app, host="0.0.0.0", port=8000)\
}