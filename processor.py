import os
import json
import logging
import tempfile
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import openai
from extrator import extrair_texto_docx

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("rubric-ai-masters")

openai.api_key = os.environ.get("OPENAI_API_KEY")

CRITERIOS = {
    "introducao": """
A Introdução deve conter OBRIGATORIAMENTE todos os elementos abaixo. Aponte como erro/melhoria qualquer ausência:

1. CONTEXTUALIZAÇÃO DO TEMA: O aluno apresenta o tema escolhido com relevância clara, preferencialmente com fundamentação teórica (citações). Aponte se falta fundamentação na contextualização.

2. PROBLEMA DE PESQUISA: Deve derivar da contextualização e OBRIGATORIAMENTE estar em forma de pergunta. Aponte como erro grave se não for uma pergunta.

3. OBJETIVO GERAL: Deve derivar diretamente do problema de pesquisa.

4. OBJETIVOS ESPECÍFICOS: Máximo de 3 ou 4. Aponte como erro se houver mais de 4.

5. JUSTIFICATIVA: Deve apresentar claramente por que a pesquisa é relevante. Aponte se estiver ausente ou vaga.

6. INDICAÇÃO METODOLÓGICA: Deve indicar brevemente qual será a metodologia, sem detalhar.

7. PARÁGRAFO DE SÍNTESE (ESTRUTURA): O último parágrafo deve descrever como o trabalho está organizado (ex: "O capítulo 2 aborda..."). Aponte como erro se estiver ausente.

8. FREQUÊNCIA DE CITAÇÕES: O texto deve conter citações/referências a cada 2 ou 3 parágrafos no mínimo. Aponte trechos longos sem citação.
""",
    "metodologia": """
A Metodologia deve conter OBRIGATORIAMENTE todos os elementos abaixo:

REGRAS GERAIS:
1. DESCRIÇÃO DA METODOLOGIA: O tipo de pesquisa (bibliográfica, de campo, etc.) deve estar claramente descrito.
2. FUNDAMENTAÇÃO DA ESCOLHA: Mínimo de 2 autores diferentes de metodologia científica fundamentando a escolha. Aponte como erro grave se houver menos de 2 autores.

SE FOR PESQUISA BIBLIOGRÁFICA (verificar todos os itens):
3. DESCRITORES/PALAVRAS-CHAVE: Quais termos foram usados na busca.
4. PERÍODO DELIMITADO: Recorte temporal definido (normalmente últimos 5 ou 10 anos).
5. PLATAFORMAS DE BUSCA: Indicação clara de onde a busca foi feita (ex: SciELO, Periódicos CAPES).
6. CRITÉRIOS DE INCLUSÃO E EXCLUSÃO: Quais regras definiram o que entra e o que sai da pesquisa.
7. DESCRIÇÃO QUANTITATIVA DO FUNIL:
   - Quantos materiais vieram inicialmente com as palavras-chave
   - Quantos foram excluídos (ex: leitura de títulos/resumos)
   - Quantos trabalhos restaram para análise final
8. QUADRO DE OBRAS RESULTANTES: Quadro apresentando as obras finais (pode estar aqui ou nos Resultados).

SE FOR PESQUISA DE CAMPO/EMPÍRICA:
9. APROVAÇÃO PRÉVIA: Deve ter sido indicada e aprovada no projeto de capstone.
10. PROCEDIMENTOS ADOTADOS: Indicação clara e detalhada dos procedimentos de coleta e análise de dados.
""",
    "fundamentacao": """
A Fundamentação Teórica deve conter OBRIGATORIAMENTE:

1. APRESENTAÇÃO DAS OBRAS: Deve apresentar e discutir as obras encontradas no levantamento bibliográfico (produções dos últimos 5 ou 10 anos).

2. ALINHAMENTO COM OBJETIVOS: O conteúdo deve contribuir para responder ao problema de pesquisa e atingir os objetivos definidos na introdução.

3. DIÁLOGO ENTRE AUTORES: Os autores devem dialogar entre si. Aponte como erro citações isoladas sem conexão entre elas.

4. ATUALIDADE DA BIBLIOGRAFIA: Aponte uso de bibliografia muito antiga (salvo obras clássicas reconhecidas).

5. FIDELIDADE AO TEMA: Aponte se o conteúdo se afasta do tema proposto nos objetivos.

6. FREQUÊNCIA DE CITAÇÕES: Citações a cada 2 ou 3 parágrafos no mínimo.
""",
    "resultados": """
Resultados e Discussão devem conter OBRIGATORIAMENTE:

1. APRESENTAÇÃO DOS RESULTADOS: Apresentação clara dos dados coletados ou das informações encontradas na literatura.

2. ANÁLISE E INTERPRETAÇÃO (DISCUSSÃO): O aluno deve interpretar, analisar e explicar o significado dos achados. Aponte se resultados são apresentados sem discussão crítica.

3. CONEXÃO COM OBJETIVOS E LITERATURA: Os achados devem ser conectados aos objetivos do trabalho e à literatura existente (diálogo com o conhecimento já produzido). Aponte se falta essa conexão.

4. QUADRO DE OBRAS RESULTANTES: Se não foi apresentado na Metodologia, DEVE estar aqui obrigatoriamente. Aponte como erro se estiver ausente em ambos os capítulos.

5. FREQUÊNCIA DE CITAÇÕES: Citações a cada 2 ou 3 parágrafos no mínimo.
"""
}


def obter_ou_criar_parte_comentarios(documento):
    """Obtém a parte de comentários do documento .docx, criando-a se necessário."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.oxml import parse_xml
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    part = documento.part
    try:
        comments_part = part.part_related_by(RT.COMMENTS)
        return comments_part.element
    except KeyError:
        pass

    comments_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    comments_element = parse_xml(comments_xml)
    partname = PackURI("/word/comments.xml")
    content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    new_part = Part(partname, content_type, comments_xml, part.package)
    new_part._element = comments_element
    part.relate_to(new_part, RT.COMMENTS)
    return comments_element


def inserir_comentario(documento, paragrafo, texto_comentario, autor="Rubric AI"):
    comentarios = obter_ou_criar_parte_comentarios(documento)
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


def extrair_resposta_json(texto_resposta: str):
    """Extrai e faz parse do JSON retornado pela IA, tolerando blocos markdown."""
    texto = texto_resposta.strip()
    if "```json" in texto:
        texto = texto.split("```json")[1].split("```")[0]
    elif "```" in texto:
        texto = texto.split("```")[1].split("```")[0]
    return json.loads(texto.strip())


async def processar_documento(caminho_versao, caminho_projeto, nome_aluno, numero_versao, capitulos):
    texto_versao = extrair_texto_docx(caminho_versao)

    contexto_projeto = "NENHUM PROJETO DE CAPSTONE FOI ENVIADO PARA ESTE ALUNO."
    if caminho_projeto:
        texto_projeto = extrair_texto_docx(caminho_projeto)
        contexto_projeto = f"""PROJETO DE CAPSTONE APROVADO (documento de referência oficial do tema, problema de pesquisa, objetivos e metodologia aprovados para este aluno):
{texto_projeto[:4000]}"""

    criterios_aplicaveis = ""
    for cap in capitulos:
        cap = cap.strip().lower()
        if cap in CRITERIOS:
            criterios_aplicaveis += f"\n=== {cap.upper()} ===\n"
            criterios_aplicaveis += CRITERIOS[cap]

    prompt_sistema = f"""Você é avaliador especialista de monografias de mestrado da Must University.
Sua função é analisar o texto enviado e gerar feedback construtivo e preciso em português brasileiro.
Norma acadêmica: APA.

{contexto_projeto}

VERIFICAÇÃO DE ADERÊNCIA AO PROJETO (CRITÉRIO OBRIGATÓRIO E PRIORITÁRIO):
Antes de qualquer outra análise, compare o conteúdo da monografia abaixo com o Projeto de Capstone acima.
- O tema tratado na monografia é o MESMO tema aprovado no projeto?
- O problema de pesquisa, os objetivos e a metodologia da monografia estão alinhados com o que foi aprovado no projeto?
- Se o aluno se afastou do tema, objetivo ou metodologia originalmente aprovados, isso é um problema GRAVE.
  Gere um comentário específico apontando claramente o desvio, citando o que foi aprovado no projeto e o que está sendo
  apresentado de diferente na monografia. Marque esse comentário com tipo "desvio_projeto".
- Se não houver projeto de capstone enviado, ignore esta verificação.

CRITÉRIOS OBRIGATÓRIOS POR CAPÍTULO:
{criterios_aplicaveis}

O TEXTO DA MONOGRAFIA ABAIXO ESTÁ NUMERADO POR PARÁGRAFO NO FORMATO "[N] texto".
Use EXATAMENTE o número N entre colchetes como "paragrafo_indice" na sua resposta. Não invente índices.

INSTRUÇÕES IMPORTANTES:
- Seja específico: aponte o problema exato e sugira como corrigir
- Use tom respeitoso e construtivo
- NÃO comente parágrafos que estão corretos
- Foque apenas em ausências, erros, melhorias necessárias e desvios em relação ao projeto aprovado
- O comentário de desvio de projeto (se houver) deve ser o primeiro a aparecer, ancorado no parágrafo mais relevante (geralmente o de contextualização ou problema de pesquisa)

Retorne APENAS um JSON válido, sem texto adicional, sem markdown:
[{{"paragrafo_indice": 0, "comentario": "texto do comentário", "tipo": "ausencia|melhoria|desvio_projeto"}}]"""

    cliente = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    resposta = cliente.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": prompt_sistema},
            {"role": "user", "content": f"Monografia de {nome_aluno} (V{numero_versao}), texto numerado por parágrafo:\n\n{texto_versao[:9000]}"}
        ],
        temperature=0.3
    )

    texto_resposta = resposta.choices[0].message.content
    logger.info(f"Resposta bruta da IA: {texto_resposta[:2000]}")

    try:
        comentarios_ia = extrair_resposta_json(texto_resposta)
    except Exception as e:
        logger.error(f"Falha ao fazer parse do JSON da IA: {e}")
        logger.error(f"Texto recebido: {texto_resposta}")
        raise ValueError(f"A IA retornou um formato inesperado e o documento não pôde ser comentado: {e}")

    if not isinstance(comentarios_ia, list):
        raise ValueError("A resposta da IA não é uma lista de comentários como esperado.")

    doc = Document(caminho_versao)
    # Índice IDÊNTICO ao usado no extrator.py: doc.paragraphs bruto, sem filtrar vazios
    paragrafos = doc.paragraphs

    inseridos = 0
    falhas = 0
    for item in comentarios_ia:
        idx = item.get("paragrafo_indice")
        comentario = item.get("comentario", "")
        tipo = item.get("tipo", "melhoria")
        if idx is None or not comentario:
            continue
        if tipo == "aprovado":
            continue
        if idx < 0 or idx >= len(paragrafos):
            logger.warning(f"Índice de parágrafo fora do intervalo: {idx} (total: {len(paragrafos)})")
            falhas += 1
            continue
        prefixo = "[DESVIO DO PROJETO]" if tipo == "desvio_projeto" else f"[Rubric AI V{numero_versao}]"
        try:
            inserir_comentario(doc, paragrafos[idx], f"{prefixo} {comentario}")
            inseridos += 1
        except Exception as e:
            logger.error(f"Falha ao inserir comentário no parágrafo {idx}: {e}")
            falhas += 1

    logger.info(f"Comentários inseridos: {inseridos} | Falhas: {falhas} | Total recebido da IA: {len(comentarios_ia)}")

    if inseridos == 0:
        raise ValueError(
            f"Nenhum comentário pôde ser inserido no documento. "
            f"A IA retornou {len(comentarios_ia)} comentário(s), mas {falhas} falharam por índice inválido."
        )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        caminho_resultado = tmp.name
    doc.save(caminho_resultado)
    return caminho_resultado
