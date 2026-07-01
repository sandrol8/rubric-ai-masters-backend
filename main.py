from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import tempfile
import os
from processor import processar_documento

app = FastAPI(title="Rubric AI Masters Backend")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "online", "servico": "Rubric AI Masters Backend"}

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/analisar")
async def analisar_documento(
    versao: UploadFile = File(...),
    projeto: UploadFile = File(None),
    nome_aluno: str = Form(...),
    numero_versao: int = Form(1),
    capitulos: str = Form("introducao,metodologia"),
    nome_professor: str = Form("Professor(a)")
):
    caminho_versao = None
    caminho_projeto = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_versao:
            conteudo_versao = await versao.read()
            tmp_versao.write(conteudo_versao)
            caminho_versao = tmp_versao.name

        if projeto:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_projeto:
                conteudo_projeto = await projeto.read()
                tmp_projeto.write(conteudo_projeto)
                caminho_projeto = tmp_projeto.name

        caminho_resultado = await processar_documento(
            caminho_versao=caminho_versao,
            caminho_projeto=caminho_projeto,
            nome_aluno=nome_aluno,
            numero_versao=numero_versao,
            capitulos=capitulos.split(","),
            nome_professor=nome_professor
        )

        nome_arquivo = f"{nome_aluno.replace(' ', '_')}_V{numero_versao}_comentado.docx"
        return FileResponse(
            caminho_resultado,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=nome_arquivo
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        if caminho_versao and os.path.exists(caminho_versao):
            os.unlink(caminho_versao)
        if caminho_projeto and os.path.exists(caminho_projeto):
            os.unlink(caminho_projeto)
