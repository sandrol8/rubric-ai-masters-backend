from docx import Document

def extrair_texto_docx(caminho: str) -> str:
    """Extrai o texto completo de um arquivo .docx."""
    try:
        doc = Document(caminho)
        paragrafos = []
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip():
                paragrafos.append(f"[{i}] {p.text.strip()}")
        return "\n".join(paragrafos)
    except Exception as e:
        return f"Erro ao extrair texto: {str(e)}"
