"""FastAPI server: dashboard UI + XML payroll API + command endpoint."""

from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from payroll_dashboard.commands import CommandError, apply_command
from payroll_dashboard.engine import build_slip, from_xml_root
from payroll_dashboard.xml_store import DEFAULT_PATH, load, root_to_dict

STATIC = Path(__file__).resolve().parent / "static"

app = FastAPI(title="Infutive Payroll Dashboard", version="1.0.0")


class CommandBody(BaseModel):
    command: str


@app.get("/api/health")
def health():
    return {"ok": True, "config": str(DEFAULT_PATH)}


@app.get("/api/payroll")
def get_payroll():
    tree = load()
    return from_xml_root(tree.getroot())


@app.get("/api/slip")
def get_slip(emp: str, month: str):
    tree = load()
    data = root_to_dict(tree.getroot())
    try:
        return build_slip(data, emp, month)
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))


@app.get("/api/xml", response_class=PlainTextResponse)
def get_xml():
    return DEFAULT_PATH.read_text(encoding="utf-8")


@app.post("/api/command")
def post_command(body: CommandBody):
    try:
        msg = apply_command(body.command)
        tree = load()
        return {"message": msg, "payroll": from_xml_root(tree.getroot())}
    except CommandError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


app.mount("/static", StaticFiles(directory=STATIC), name="static")
