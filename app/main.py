from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from app.api.parser import router as parser_router
import os

app = FastAPI(title="AccuSync - Bank Statement Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(parser_router, prefix="/api/v1")

@app.get("/", response_class=HTMLResponse)
def home_ui():
    template_path = os.path.join("templates", "dashboard.html")
    with open(template_path, "r", encoding="utf-8") as f:
        return f.read()

@app.get("/dashboard", response_class=HTMLResponse)
def dashboard_ui():
    template_path = os.path.join("templates", "dashboard.html")
    with open(template_path, "r", encoding="utf-8") as f:
        return f.read()