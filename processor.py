import os
import tempfile
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import openai
from extrator import extrair_texto_docx

openai.api_key = os.environ.get("OPENAI_API_KEY")

CRITERIOS = {
    "introducao": """
A Introdução deve conter obrigatoriamente:
1. Contextualização do tema com relevância clara
2. Problema de pesquisa (SEMPRE em forma de pergunta)
3. Objetivo geral
4. Até 4 objetivos específicos
5. Justificativa
6. Indicação da metodologia (sem detalhar)
7. Parágrafo final descrevendo os capítulos
8. Referências a cada 2-3 parágrafos (mínimo)
""",
    "metodologia": """
A Metodologia deve conter obrigatoriamente:
1. Descrição detalhada da metodologia adotada
2. Mínimo 2 autores fundamentando a escolha
3. Se bibliográfica: descritores, período, plataformas, quadro de materiais, critérios de inclusão/exclusão
4. Se campo: procedimentos detalhados
""",
    "fundamentacao": """
A Fundamentação Teórica deve conter:
1. Apresentação das obras do levantamento
2. Resposta ao problema de pesquisa
""",
    "resultados": """
Resultados e Discussão devem conter:
1. Apresentação dos dados
2. Interpretação e análise crítica
3. Conexão com os objetivos
4. Diálogo com a literatura
"""
}


def inserir_comentario(paragrafo, texto_comentario, autor="Rubric AI"):
    doc = paragrafo._element.getparent().getparent().getparent()
    comentarios = doc.find(qn('w:comments'))
    if comentarios is None:
        comentarios = OxmlElement('w:comments')
        doc.append(comentarios)
    ids_existentes = [int(c.get(qn('w:id'), 0)) for c in comentarios.findall(qn('w:comment'))]
    novo_id = max(ids_existentes, default=0) + 1
    comentario = OxmlElement('w:comment')
    comentario.set(qn('w:id'), str(novo_id))
    comentario.set(qn('w:author'), autor)
    comentario.set(qn('w:date'), datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"))
    comentario.set(qn('w:initials'), "RA")
    p_comentario = OxmlElement('w:p')
    r_comentario = OxmlElement('w:r')
    t_comentario = OxmlElement('w:t')
    t_comentario.text = texto_comentario
    t_comentario.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r_comentario.append(t_comentario)
    p_comentario.append(r_comentario)
    comentario.append(p_comentario)
    comentarios.append(comentario)
    r_start = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'CommentReference')
    rpr.append(rStyle)
    r_start.append(rpr)
    commentRef = OxmlElement('w:commentReference')
    commentRef.set(qn('w:id'), str(novo_id))
    r_start.append(commentRef)
    paragrafo._element.append(r_start)


async def processar_documento(caminho_versao, caminho_projeto, nome_aluno, numero_versao, capitulos):
    from extrator import extrair_texto_docx
    texto_versao = extrair_texto_docx(caminho_versao)
    contexto_projeto = ""
    if caminho_projeto:
        texto_projeto = extrair_texto_docx(caminho_projeto)
        contexto_projeto = f"PROJETO DE CAPSTONE:\n{texto_projeto[:3000]}"
    criterios_aplicaveis = ""
    for cap in capitulos:
        cap = cap.strip().lower()
        if cap in CRITERIOS:
            criterios_aplicaveis += CRITERIOS[cap]
    prompt_sistema = f"""Você é avaliador de monografias de mestrado da Must University.
Analise o texto e gere feedback construtivo em português brasileiro. Norma: APA.
{contexto_projeto}
CRITÉRIOS:
{criterios_aplicaveis}
Retorne APENAS JSON:
[{{"paragrafo_indice": 0, "comentario": "texto", "tipo": "ausencia|melhoria|aprovado"}}]"""
    cliente = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    resposta = cliente.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": prompt_sistema},
            {"role": "user", "content": f"Monografia de {nome_aluno} (V{numero_versao}):\n\n{texto_versao[:8000]}"}
        ],
        temperature=0.3
    )
    import json
    texto_resposta = resposta.choices[0].message.content.strip()
    if "```json" in texto_resposta:
        texto_resposta = texto_resposta.split("```json")[1].split("```")[0]
    elif "```" in texto_resposta:
        texto_resposta = texto_resposta.split("```")[1].split("```")[0]
    comentarios_ia = json.loads(texto_resposta)
    doc = Document(caminho_versao)
    paragrafos = [p for p in doc.paragraphs if p.text.strip()]
    for item in comentarios_ia:
        idx = item.get("paragrafo_indice", 0)
        comentario = item.get("comentario", "")
        tipo = item.get("tipo", "melhoria")
        if tipo != "aprovado" and idx < len(paragrafos):
            try:
                inserir_comentario(paragrafos[idx], f"[Rubric AI V{numero_versao}] {comentario}")
            except Exception:
                pass
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        caminho_resultado = tmp.name
    doc.save(caminho_resultado)
    return caminho_resultado
