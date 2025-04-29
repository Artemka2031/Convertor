from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from tempfile import NamedTemporaryFile
from pathlib import Path
from scripts.excel_to_xml import excel_to_xml
from scripts.xml_to_excel import xml_to_excel

app = FastAPI(title="UTD Converter")

WEB_DIR = Path(__file__).resolve().parent.parent / "web"

def _tmp(suffix: str) -> Path:
    return Path(NamedTemporaryFile(delete=False, suffix=suffix).name)

# ---------- API ----------------------------------------------------------------
@app.post("/excel-to-xml", response_class=FileResponse)
async def excel_to_xml_endpoint(file: UploadFile):
    if not file.filename.lower().endswith((".xls", ".xlsx")):
        raise HTTPException(400, "Нужен Excel (.xls/.xlsx)")
    src, dst = _tmp(".xlsx"), _tmp(".xml")
    src.write_bytes(await file.read())

    excel_to_xml(src, dst)
    return FileResponse(dst, media_type="application/xml",
                        filename=Path(file.filename).with_suffix(".xml").name)

@app.post("/xml-to-excel", response_class=FileResponse)
async def xml_to_excel_endpoint(file: UploadFile):
    if not file.filename.lower().endswith(".xml"):
        raise HTTPException(400, "Нужен XML")
    src, dst = _tmp(".xml"), _tmp(".xlsx")
    src.write_bytes(await file.read())

    xml_to_excel(src, dst)
    return FileResponse(dst,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=Path(file.filename).with_suffix(".xlsx").name)

# ---------- UI -----------------------------------------------------------------
@app.get("/", response_class=FileResponse)
async def index():
    return WEB_DIR / "index.html"

@app.get("/favicon.ico", response_class=Response)
async def favicon():
    # возвращаем 204, чтобы не было 404 в логах
    return Response(status_code=204)

# ---------- статика ------------------------------------------------------------
app.mount("/static", StaticFiles(directory=WEB_DIR, html=False), name="static")
