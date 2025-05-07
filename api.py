from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import subprocess
import os
from pathlib import Path
import shutil
import uuid
from typing import Optional
import json

# Load configuration
def load_config():
    try:
        with open('config.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {
            "LIBREOFFICE_PATH": r"C:\Program Files\LibreOffice\program\soffice.exe",
            "UPLOAD_DIR": "temp_uploads",
            "OUTPUT_DIR": "temp_output"
        }

config = load_config()

# Create necessary directories
Path(config["UPLOAD_DIR"]).mkdir(exist_ok=True)
Path(config["OUTPUT_DIR"]).mkdir(exist_ok=True)

app = FastAPI(
    title="PPTX to PDF Converter API",
    description="API service to convert PowerPoint (PPTX) files to PDF format",
    version="1.0.0"
)

# Mount static files directory (for css, js)
app.mount("/static", StaticFiles(directory="static"), name="static")

async def convert_to_pdf(input_file: str, output_dir: str) -> Optional[str]:
    try:
        cmd = [
            config["LIBREOFFICE_PATH"],
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            input_file
        ]
        
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        
        stdout, stderr = process.communicate()
        
        if process.returncode == 0:
            pdf_filename = Path(input_file).stem + ".pdf"
            pdf_path = os.path.join(output_dir, pdf_filename)
            if os.path.exists(pdf_path):
                return pdf_path
        
        error_msg = stderr.decode(errors='ignore').strip()
        if not error_msg:
            error_msg = stdout.decode(errors='ignore').strip()
        raise HTTPException(status_code=500, detail=f"Conversion failed: {error_msg}")
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    return None

@app.post("/convert/", response_class=FileResponse)
async def convert_pptx_to_pdf(file: UploadFile = File(...)):
    # Validate file extension
    if not file.filename.lower().endswith('.pptx'):
        raise HTTPException(status_code=400, detail="Only PPTX files are accepted")
    
    # Create unique filename to prevent conflicts
    unique_id = str(uuid.uuid4())
    temp_input_path = os.path.join(config["UPLOAD_DIR"], f"{unique_id}_{file.filename}")
    
    try:
        # Save uploaded file
        with open(temp_input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Convert to PDF
        pdf_path = await convert_to_pdf(temp_input_path, config["OUTPUT_DIR"])
        
        if pdf_path and os.path.exists(pdf_path):
            # Return the PDF file
            return FileResponse(
                pdf_path,
                media_type="application/pdf",
                filename=f"{Path(file.filename).stem}.pdf",
                background=None  # Ensures file is deleted after sending
            )
        else:
            raise HTTPException(status_code=500, detail="PDF conversion failed")
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
        
    finally:
        # Cleanup temporary files
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)

@app.get("/health")
async def health_check():
    """Health check endpoint to verify if the service is running"""
    return {"status": "healthy", "message": "Service is running"}

# Serve the HTML page
@app.get("/", response_class=HTMLResponse)
async def read_root():
    try:
        with open("index.html", "r") as f:
            html_content = f.read()
        return HTMLResponse(content=html_content, status_code=200)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="index.html not found")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 