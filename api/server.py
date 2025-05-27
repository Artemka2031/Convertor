import os
import shutil
import tempfile
import sys

# Добавляем корень проекта в sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from utils import setup_logging
from excel_to_xml import excel_to_xml
from xml_to_excel import xml_to_excel

logger = setup_logging()

app = FastAPI()

app.mount("/static", StaticFiles(directory="web"), name="static")

# Создаём папку uploads, если её нет
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)


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
        # Сохраняем входной файл в папку uploads
        input_path = os.path.join(UPLOAD_DIR, f"input_{file.filename}")
        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        logger.debug(f"Сохранён входной файл XML: {input_path}")

        # Создаём путь для выходного файла
        output_filename = f"output_{file.filename.rsplit('.', 1)[0]}.xlsx"
        output_path = os.path.join(UPLOAD_DIR, output_filename)

        # Выполняем конвертацию
        xml_to_excel(input_path, output_path)
        logger.info(f"Файл успешно конвертирован в Excel: {output_path}")

        # Отправляем файл клиенту
        response = FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_filename
        )
        return response
    except Exception:
        raise
        # logger.error(f"Ошибка при конвертации XML в Excel: {e}")
        # return {"error": str(e)}
    finally:
        # Удаляем входной файл после обработки
        if 'input_path' in locals() and os.path.exists(input_path):
            os.unlink(input_path)
        # Выходной файл удаляем только после отправки, поэтому оставляем его до завершения запроса


@app.post("/excel-to-xml")
async def convert_excel_to_xml(file: UploadFile = File(...)):
    logger.debug("Получен запрос на конвертацию Excel в XML")
    try:
        # Сохраняем входной файл в папку uploads
        input_path = os.path.join(UPLOAD_DIR, f"input_{file.filename}")
        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        logger.debug(f"Сохранён входной файл Excel: {input_path}")

        # Создаём путь для выходного файла
        output_filename = f"output_{file.filename.rsplit('.', 1)[0]}.xml"
        output_path = os.path.join(UPLOAD_DIR, output_filename)

        # Выполняем конвертацию
        excel_to_xml(input_path, output_path)
        logger.info(f"Файл успешно конвертирован в XML: {output_path}")

        # Отправляем файл клиенту
        response = FileResponse(
            output_path,
            media_type="application/xml",
            filename=output_filename
        )
        return response
    except Exception as e:
        raise
        # logger.error(f"Ошибка при конвертации Excel в XML: {e}")
        # return {"error": str(e)}
    finally:
        # Удаляем входной файл после обработки
        if 'input_path' in locals() and os.path.exists(input_path):
            os.unlink(input_path)
        # Выходной файл удаляем только после отправки, поэтому оставляем его до завершения запроса


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)