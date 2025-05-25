import os
import shutil
import tempfile

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

# Заменяем на абсолютные импорты
import excel_to_xml
import xml_to_excel
from utils import setup_logging

# Остальной код остаётся без изменений
logger = setup_logging()

app = FastAPI()

app.mount("/static", StaticFiles(directory="web"), name="static")


@app.get("/", response_class=HTMLResponse)
async def read_root():
    logger.debug("Запрос главной страницы")
    with open("web/index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/health")
async def health_check():
    return {"status": "healthy"}


@app.post("/xml-to-excel")
async def convert_xml_to_excel(file: UploadFile = File(...)):
    logger.debug("Получен запрос на конвертацию XML в Excel")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_file:
            shutil.copyfileobj(file.file, tmp_file)
            tmp_file_path = tmp_file.name
        logger.debug(f"Сохранён временный файл XML: {tmp_file_path}")

        output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        xml_to_excel.xml_to_excel(tmp_file_path, output_file)
        logger.info(f"Файл успешно конвертирован в Excel: {output_file}")

        return FileResponse(output_file, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename="output.xlsx")
    except Exception as e:
        logger.error(f"Ошибка при конвертации XML в Excel: {e}")
        return {"error": str(e)}
    finally:
        if 'tmp_file_path' in locals() and os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path)
        if 'output_file' in locals() and os.path.exists(output_file):
            os.unlink(output_file)


@app.post("/excel-to-xml")
async def convert_excel_to_xml(file: UploadFile = File(...)):
    logger.debug("Получен запрос на конвертацию Excel в XML")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            shutil.copyfileobj(file.file, tmp_file)
            tmp_file_path = tmp_file.name
        logger.debug(f"Сохранён временный файл Excel: {tmp_file_path}")

        output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xml").name
        excel_to_xml.excel_to_xml(tmp_file_path, output_file)
        logger.info(f"Файл успешно конвертирован в XML: {output_file}")

        return FileResponse(output_file, media_type="application/xml", filename="output.xml")
    except Exception as e:
        logger.error(f"Ошибка при конвертации Excel в XML: {e}")
        return {"error": str(e)}
    finally:
        if 'tmp_file_path' in locals() and os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path)
        if 'output_file' in locals() and os.path.exists(output_file):
            os.unlink(output_file)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)