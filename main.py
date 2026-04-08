from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from docx import Document
import io
from pathlib import Path

app = FastAPI()

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@app.get("/")
async def main():
    index_path = Path(__file__).resolve().parent / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=404, detail="index.html não encontrado")
    return FileResponse(index_path)


@app.post("/edit-metadata/")
async def edit_metadata(
    file: UploadFile = File(...),
    title: str = Form(None),
    creator: str = Form(None),
    keyword: str = Form(None),
    description: str = Form(None),
    category: str = Form(None),
):
    # valida tipo
    if file.content_type != DOCX_MIME:
        raise HTTPException(status_code=400, detail="Arquivo deve ser .docx")

    content = await file.read()

    try:
        doc = Document(io.BytesIO(content))
    except Exception:
        raise HTTPException(status_code=400, detail="Arquivo DOCX inválido")

    props = doc.core_properties

    if title is not None:
        props.title = title

    if creator is not None:
        props.author = creator

    if keyword is not None:
        props.keywords = keyword

    if description is not None:
        props.comments = description  # limitação da lib

    if category is not None:
        props.category = category

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    safe_name = Path(file.filename).name
    filename = f"editado_{safe_name}"

    return StreamingResponse(
        output,
        media_type=DOCX_MIME,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        },
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)