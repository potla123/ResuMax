import os
import io
import re
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv

import resume_optimizer

# Load environment variables from .env if present
load_dotenv()

app = FastAPI(title="AI Resume Creator & Optimizer")

# Add CORS middleware to allow localhost calls
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/scrape-jd")
def scrape_jd(url: str = Form(...)):
    """API endpoint to scrape job description text from a URL."""
    try:
        text = resume_optimizer.scrape_job_description(url)
        return {"text": text}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/optimize")
async def optimize_resume(
    resume: UploadFile = File(...),
    jd_text: str = Form(None),
    jd_url: str = Form(None),
    openai_key: str = Form(None),
    anthropic_key: str = Form(None)
):
    """API endpoint to run the full resume optimization pipeline."""
    # 1. Resolve keys (Form data or .env)
    openai_api_key = openai_key or os.environ.get("OPENAI_API_KEY")
    anthropic_api_key = anthropic_key or os.environ.get("ANTHROPIC_API_KEY")
    
    if not openai_api_key:
        raise HTTPException(
            status_code=400,
            detail="OpenAI API key is missing. Please configure it in your Settings or .env file."
        )
    if not anthropic_api_key:
        raise HTTPException(
            status_code=400,
            detail="Anthropic API key is missing. Please configure it in your Settings or .env file."
        )

    # 2. Extract Resume Text
    filename = resume.filename.lower()
    content_bytes = await resume.read()
    
    if filename.endswith(".pdf"):
        try:
            resume_text = resume_optimizer.extract_text_from_pdf(io.BytesIO(content_bytes))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Failed to read PDF: {str(e)}")
    elif filename.endswith(".docx"):
        try:
            resume_text = resume_optimizer.extract_text_from_docx(io.BytesIO(content_bytes))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Failed to read Word file: {str(e)}")
    elif filename.endswith(".doc"):
        raise HTTPException(
            status_code=400,
            detail="Legacy .doc files aren't supported. Please open your resume in Word and use 'Save As' to save it as .docx, then re-upload."
        )
    elif filename.endswith(".txt"):
        try:
            resume_text = content_bytes.decode("utf-8")
        except UnicodeDecodeError:
            try:
                resume_text = content_bytes.decode("latin-1")
            except Exception as e:
                raise HTTPException(status_code=400, detail=f"Failed to decode text file: {str(e)}")
    else:
        raise HTTPException(
            status_code=400,
            detail="Unsupported file format. Please upload a PDF, DOCX, or TXT file."
        )

    if not resume_text.strip():
        raise HTTPException(status_code=400, detail="The uploaded resume is empty.")

    # 3. Resolve Job Description
    resolved_jd = ""
    if jd_text and jd_text.strip():
        resolved_jd = jd_text.strip()
    elif jd_url and jd_url.strip():
        try:
            resolved_jd = resume_optimizer.scrape_job_description(jd_url.strip())
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Failed to scrape job description: {str(e)}")
    
    if not resolved_jd.strip():
        raise HTTPException(
            status_code=400,
            detail="Job description is required. Please paste text or provide a valid URL."
        )

    # 4. Run Pipeline
    try:
        results = resume_optimizer.run_optimization_pipeline(
            openai_key=openai_api_key,
            anthropic_key=anthropic_api_key,
            resume_text=resume_text,
            jd_text=resolved_jd
        )
        return results
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Pipeline error: {str(e)}")

class DownloadPayload(BaseModel):
    content: str
    filename: str = "Resume"

_SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9_.\- ]+")

def _safe_filename(name: str, ext: str) -> str:
    base = _SAFE_NAME_RE.sub("", (name or "Resume")).strip().replace(" ", "_") or "Resume"
    return f"{base}.{ext}"

@app.post("/api/download/docx")
def download_docx(payload: DownloadPayload):
    if not payload.content.strip():
        raise HTTPException(status_code=400, detail="No content provided.")
    try:
        data = resume_optimizer.markdown_to_docx_bytes(payload.content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to render Word document: {str(e)}")
    headers = {"Content-Disposition": f'attachment; filename="{_safe_filename(payload.filename, "docx")}"'}
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )

@app.post("/api/download/pdf")
def download_pdf(payload: DownloadPayload):
    if not payload.content.strip():
        raise HTTPException(status_code=400, detail="No content provided.")
    try:
        data = resume_optimizer.markdown_to_pdf_bytes(payload.content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to render PDF: {str(e)}")
    headers = {"Content-Disposition": f'attachment; filename="{_safe_filename(payload.filename, "pdf")}"'}
    return StreamingResponse(io.BytesIO(data), media_type="application/pdf", headers=headers)

@app.post("/api/download/txt")
def download_txt(payload: DownloadPayload):
    if not payload.content.strip():
        raise HTTPException(status_code=400, detail="No content provided.")
    text = resume_optimizer.markdown_to_plaintext(payload.content)
    headers = {"Content-Disposition": f'attachment; filename="{_safe_filename(payload.filename, "txt")}"'}
    return StreamingResponse(io.BytesIO(text.encode("utf-8")), media_type="text/plain; charset=utf-8", headers=headers)

# Ensure static folder exists
os.makedirs("static", exist_ok=True)

# Mount the static directory to serve HTML/CSS/JS frontend
app.mount("/", StaticFiles(directory="static", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    # Start the server on port 8000
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
