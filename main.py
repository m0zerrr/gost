from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
import uuid
import uvicorn
from gost_rules import format_to_gost, split_tables_only

app = FastAPI()
app.mount('/static', StaticFiles(directory='static'), name='static')

def cleanup_file(path):
    if os.path.exists(path):
        os.remove(path)

@app.get('/', response_class=HTMLResponse)
async def read_index():
    with open('static/index.html', 'r', encoding='utf-8') as f:
        return f.read()

@app.post('/format-text')
async def format_text(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Этап 1: оформление текста БЕЗ разбиения таблиц."""
    uid = str(uuid.uuid4())
    input_path = f'temp_{uid}_input.docx'
    output_path = f'temp_{uid}_formatted.docx'
    with open(input_path, 'wb') as buffer:
        content = await file.read()
        buffer.write(content)
    try:
        format_to_gost(input_path, output_path)
    except Exception as e:
        cleanup_file(input_path)
        return {'error': f'Ошибка: {str(e)}'}
    cleanup_file(input_path)
    background_tasks.add_task(cleanup_file, output_path)
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f'formatted_{file.filename}'
    )

@app.post('/split-tables')
async def split_tables(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Этап 2: разбиение таблиц после ручного добавления разрывов."""
    uid = str(uuid.uuid4())
    input_path = f'temp_{uid}_input.docx'
    output_path = f'temp_{uid}_final.docx'
    with open(input_path, 'wb') as buffer:
        content = await file.read()
        buffer.write(content)
    try:
        split_tables_only(input_path, output_path)
    except Exception as e:
        cleanup_file(input_path)
        return {'error': f'Ошибка: {str(e)}'}
    cleanup_file(input_path)
    background_tasks.add_task(cleanup_file, output_path)
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f'GOST_{file.filename}'
    )