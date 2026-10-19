from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
import uuid
import uvicorn
from gost_rules import format_to_gost

app = FastAPI()

app.mount('/static', StaticFiles(directory='static'), name='static')

def cleanup_file(path):
  if os.path.exists(path):
    os.remove(path)

@app.get('/', response_class=HTMLResponse)
async def read_index():
  with open('static/index.html', 'r', encoding='utf-8') as f:
    return f.read()
  
@app.post('/format')
async def format_file(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
  uid = str(uuid.uuid4())
  input_path = f'temp_{uid}_input.docx'
  output_path = f'temp_{uid}_output.docx'
  
  with open(input_path, 'wb') as buffer:
    content = await file.read()
    buffer.write(content)
    
  try:
    format_to_gost(input_path, output_path)
  except Exception as e:
    cleanup_file(input_path)
    return {'error': f'Ошибка при обработке файла: {str(e)}'}
  
  cleanup_file(input_path)
  background_tasks.add_task(cleanup_file, output_path)
  
  return FileResponse(
    output_path,
    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    filename=f'GOST_{file.filename}'
  )
