import os
import json
import logging
import shutil
import tempfile
import zipfile
from lxml import etree
from datetime import datetime
import openai
from extrator import extrair_texto_docx

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("rubric-ai-masters")

openai.api_key = os.environ.get("OPENAI_API_KEY")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NSMAP = {"w": W_NS}

CRITERIOS = {
    "introducao": """
A Introducao deve conter OBRIGATORIAMENTE todos os elementos abaixo. Aponte como erro/melhoria qualquer ausencia:

1. CONTEXTUALIZACAO DO TEMA: O aluno apresenta o tema escolhido com relevancia clara, preferencialmente com fundamentacao teorica (citacoes). Aponte se falta fundamentacao na contextualizacao.

2. PROBLEMA DE PESQUISA: Deve derivar da contextualizacao e OBRIGATORIAMENTE estar em forma de pergunta. Aponte como erro grave se nao for uma pergunta.

3. OBJETIVO GERAL: Deve derivar diretamente do problema de pesquisa.

4. OBJETIVOS ESPECIFICOS: Maximo de 3 ou 4. Aponte como erro se houver mais de 4.

5. JUSTIFICATIVA: Deve apresentar claramente por que a pesquisa e relevante. Aponte se estiver ausente ou vaga.

6. INDICACAO METODOLOGICA: Deve indicar brevemente qual sera a metodologia, sem detalhar.

7. PARAGRAFO DE SINTESE (ESTRUTURA): O ultimo paragrafo deve descrever como o trabalho esta organizado (ex: "O capitulo 2 aborda..."). Aponte como erro se estiver ausente.

8. FREQUENCIA DE CITACOES: O texto deve conter citacoes/referencias a cada 2 ou 3 paragrafos no minimo. Aponte trechos longos sem citacao.
""",
    "metodologia": """
A Metodologia deve conter OBRIGATORIAMENTE todos os elementos abaixo:

REGRAS GERAIS:
1. DESCRICAO DA METODOLOGIA: O tipo de pesquisa (bibliografica, de campo, etc.) deve estar claramente descrito.
2. FUNDAMENTACAO DA ESCOLHA: Minimo de 2 autores diferentes de metodologia cientifica fundamentando a escolha. Aponte como erro grave se houver menos de 2 autores.

SE FOR PESQUISA BIBLIOGRAFICA (verificar todos os itens):
3. DESCRITORES/PALAVRAS-CHAVE: Quais termos foram usados na busca.
4. PERIODO DELIMITADO: Recorte temporal definido (normalmente ultimos 5 ou 10 anos).
5. PLATAFORMAS DE BUSCA: Indicacao clara de onde a busca foi feita (ex: SciELO, Periodicos CAPES).
6. CRITERIOS DE INCLUSAO E EXCLUSAO: Quais regras definiram o que entra e o que sai da pesquisa.
7. DESCRICAO QUANTITATIVA DO FUNIL:
   - Quantos materiais vieram inicialmente com as palavras-chave
   - Quantos foram excluidos (ex: leitura de titulos/resumos)
   - Quantos trabalhos restaram para analise final
8. QUADRO DE OBRAS RESULTANTES: Quadro apresentando as obras finais (pode estar aqui ou nos Resultados).

SE FOR PESQUISA DE CAMPO/EMPIRICA:
9. APROVACAO PREVIA: Deve ter sido indicada e aprovada no projeto de capstone.
10. PROCEDIMENTOS ADOTADOS: Indicacao clara e detalhada dos procedimentos de coleta e analise de dados.
""",
    "fundamentacao": """
A Fundamentacao Teorica deve conter OBRIGATORIAMENTE:

1. APRESENTACAO DAS OBRAS: Deve apresentar e discutir as obras encontradas no levantamento bibliografico (producoes dos ultimos 5 ou 10 anos).

2. ALINHAMENTO COM OBJETIVOS: O conteudo deve contribuir para responder ao problema de pesquisa e atingir os objetivos definidos na introducao.

3. DIALOGO ENTRE AUTORES: Os autores devem dialogar entre si. Aponte como erro citacoes isoladas sem conexao entre elas.

4. ATUALIDADE DA BIBLIOGRAFIA: Aponte uso de bibliografia muito antiga (salvo obras classicas reconhecidas).

5. FIDELIDADE AO TEMA: Aponte se o conteudo se afasta do tema proposto nos objetivos.

6. FREQUENCIA DE CITACOES: Citacoes a cada 2 ou 3 paragrafos no minimo.
""",
    "resultados": """
Resultados e Discussao devem conter OBRIGATORIAMENTE:

1. APRESENTACAO DOS RESULTADOS: Apresentacao clara dos dados coletados ou das informacoes encontradas na literatura.

2. ANALISE E INTERPRETACAO (DISCUSSAO): O aluno deve interpretar, analisar e explicar o significado dos achados. Aponte se resultados sao apresentados sem discussao critica.

3. CONEXAO COM OBJETIVOS E LITERATURA: Os achados devem ser conectados aos objetivos do trabalho e a literatura existente (dialogo com o conhecimento ja produzido). Aponte se falta essa conexao.

4. QUADRO DE OBRAS RESULTANTES: Se nao foi apresentado na Metodologia, DEVE estar aqui obrigatoriamente. Aponte como erro se estiver ausente em ambos os capitulos.

5. FREQUENCIA DE CITACOES: Citacoes a cada 2 ou 3 paragrafos no minimo.
"""
}


def extrair_resposta_json(texto_resposta: str):
    """Extrai e faz parse do JSON retornado pela IA, tolerando blocos markdown."""
    texto = texto_resposta.strip()
    if "```json" in texto:
        texto = texto.split("```json")[1].split("```")[0]
    elif "```" in texto:
        texto = texto.split("```")[1].split("```")[0]
    return json.loads(texto.strip())


def _gerar_iniciais(nome):
    partes = [p for p in nome.replace(".", "").strip().split() if p]
    if not partes:
        return "PR"
    if len(partes) == 1:
        return partes[0][:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


def _qn(tag):
    return "{%s}%s" % (W_NS, tag)


def inserir_comentarios_no_docx(caminho_entrada, caminho_saida, comentarios_por_paragrafo, numero_versao, nome_professor="Professor(a)"):
    """
    Insere comentarios nativos do Word manipulando diretamente o pacote ZIP do .docx.
    comentarios_por_paragrafo: lista de tuplas (indice_paragrafo, texto_comentario, tipo)
    Retorna (inseridos, falhas).
    """
    pasta_temp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(caminho_entrada, 'r') as z:
            z.extractall(pasta_temp)

        document_xml_path = os.path.join(pasta_temp, "word", "document.xml")
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(document_xml_path, parser)
        root = tree.getroot()
        body = root.find(_qn("body"))
        paragrafos_xml = body.findall(_qn("p"))

        comments_xml_path = os.path.join(pasta_temp, "word", "comments.xml")
        if os.path.exists(comments_xml_path):
            comments_tree = etree.parse(comments_xml_path, parser)
            comments_root = comments_tree.getroot()
        else:
            comments_root = etree.Element(_qn("comments"), nsmap=NSMAP)

        ids_existentes = [int(c.get(_qn("id"), 0)) for c in comments_root.findall(_qn("comment"))]
        proximo_id = max(ids_existentes, default=-1) + 1

        inseridos = 0
        falhas = 0

        for idx, texto_comentario, tipo in comentarios_por_paragrafo:
            if idx < 0 or idx >= len(paragrafos_xml):
                logger.warning("Indice de paragrafo fora do intervalo: %s (total: %s)" % (idx, len(paragrafos_xml)))
                falhas += 1
                continue
            try:
                paragrafo_xml = paragrafos_xml[idx]
                comment_id = str(proximo_id)
                proximo_id += 1

                comment_el = etree.SubElement(comments_root, _qn("comment"))
                comment_el.set(_qn("id"), comment_id)
                comment_el.set(_qn("author"), nome_professor)
                comment_el.set(_qn("date"), datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"))
                comment_el.set(_qn("initials"), _gerar_iniciais(nome_professor))
                p_el = etree.SubElement(comment_el, _qn("p"))
                r_el = etree.SubElement(p_el, _qn("r"))
                t_el = etree.SubElement(r_el, _qn("t"))
                t_el.text = texto_comentario
                t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

                ref_start = etree.SubElement(paragrafo_xml, _qn("commentRangeStart"))
                ref_start.set(_qn("id"), comment_id)
                ref_end = etree.SubElement(paragrafo_xml, _qn("commentRangeEnd"))
                ref_end.set(_qn("id"), comment_id)
                run_ref = etree.SubElement(paragrafo_xml, _qn("r"))
                rpr_ref = etree.SubElement(run_ref, _qn("rPr"))
                rstyle_ref = etree.SubElement(rpr_ref, _qn("rStyle"))
                rstyle_ref.set(_qn("val"), "CommentReference")
                comment_ref = etree.SubElement(run_ref, _qn("commentReference"))
                comment_ref.set(_qn("id"), comment_id)

                inseridos += 1
            except Exception as e:
                logger.error("Falha ao inserir comentario no paragrafo %s: %s" % (idx, e))
                falhas += 1

        if inseridos == 0:
            return 0, falhas

        os.makedirs(os.path.dirname(comments_xml_path), exist_ok=True)
        comments_tree_final = etree.ElementTree(comments_root)
        comments_tree_final.write(comments_xml_path, xml_declaration=True, encoding="UTF-8", standalone=True)

        tree.write(document_xml_path, xml_declaration=True, encoding="UTF-8", standalone=True)

        _garantir_relacionamento_comentarios(pasta_temp)
        _garantir_content_type_comentarios(pasta_temp)

        if os.path.exists(caminho_saida):
            os.remove(caminho_saida)
        with zipfile.ZipFile(caminho_saida, 'w', zipfile.ZIP_DEFLATED) as zf:
            for raiz, _, arquivos in os.walk(pasta_temp):
                for nome_arquivo in arquivos:
                    caminho_completo = os.path.join(raiz, nome_arquivo)
                    caminho_relativo = os.path.relpath(caminho_completo, pasta_temp)
                    zf.write(caminho_completo, caminho_relativo)

        return inseridos, falhas
    finally:
        shutil.rmtree(pasta_temp, ignore_errors=True)


def _garantir_relacionamento_comentarios(pasta_temp):
    rels_path = os.path.join(pasta_temp, "word", "_rels", "document.xml.rels")
    parser = etree.XMLParser(remove_blank_text=False)
    if os.path.exists(rels_path):
        tree = etree.parse(rels_path, parser)
        root = tree.getroot()
    else:
        os.makedirs(os.path.dirname(rels_path), exist_ok=True)
        root = etree.Element("{%s}Relationships" % RELS_NS, nsmap={None: RELS_NS})
        tree = etree.ElementTree(root)

    ja_existe = any(r.get("Type", "").endswith("/comments") for r in root)
    if not ja_existe:
        ids = [r.get("Id", "") for r in root]
        numeros = [int(i.replace("rId", "")) for i in ids if i.startswith("rId") and i.replace("rId", "").isdigit()]
        novo_id = "rId%s" % (max(numeros, default=0) + 1)
        rel = etree.SubElement(root, "{%s}Relationship" % RELS_NS)
        rel.set("Id", novo_id)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
        rel.set("Target", "comments.xml")

    tree.write(rels_path, xml_declaration=True, encoding="UTF-8", standalone=True)


def _garantir_content_type_comentarios(pasta_temp):
    ct_path = os.path.join(pasta_temp, "[Content_Types].xml")
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(ct_path, parser)
    root = tree.getroot()
    ja_existe = any(
        el.get("PartName") == "/word/comments.xml" for el in root
        if el.tag == "{%s}Override" % CT_NS
    )
    if not ja_existe:
        override = etree.SubElement(root, "{%s}Override" % CT_NS)
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    tree.write(ct_path, xml_declaration=True, encoding="UTF-8", standalone=True)


async def processar_documento(caminho_versao, caminho_projeto, nome_aluno, numero_versao, capitulos, nome_professor="Professor(a)"):
    texto_versao = extrair_texto_docx(caminho_versao)

    contexto_projeto = "NENHUM PROJETO DE CAPSTONE FOI ENVIADO PARA ESTE ALUNO."
    if caminho_projeto:
        texto_projeto = extrair_texto_docx(caminho_projeto)
        contexto_projeto = "PROJETO DE CAPSTONE APROVADO (documento de referencia oficial do tema, problema de pesquisa, objetivos e metodologia aprovados para este aluno):\n%s" % texto_projeto[:4000]

    criterios_aplicaveis = ""
    for cap in capitulos:
        cap = cap.strip().lower()
        if cap in CRITERIOS:
            criterios_aplicaveis += "\n=== %s ===\n" % cap.upper()
            criterios_aplicaveis += CRITERIOS[cap]

    prompt_sistema = """Voce e avaliador especialista de monografias de mestrado da Must University.
Sua funcao e analisar o texto enviado e gerar feedback construtivo e preciso em portugues brasileiro.
Norma academica: APA.

%s

VERIFICACAO DE ADERENCIA AO PROJETO (CRITERIO OBRIGATORIO E PRIORITARIO):
Antes de qualquer outra analise, compare o conteudo da monografia abaixo com o Projeto de Capstone acima.
- O tema tratado na monografia e o MESMO tema aprovado no projeto?
- O problema de pesquisa, os objetivos e a metodologia da monografia estao alinhados com o que foi aprovado no projeto?
- Se o aluno se afastou do tema, objetivo ou metodologia originalmente aprovados, isso e um problema GRAVE.
  Gere um comentario especifico apontando claramente o desvio, citando o que foi aprovado no projeto e o que esta sendo
  apresentado de diferente na monografia. Marque esse comentario com tipo "desvio_projeto".
- Se nao houver projeto de capstone enviado, ignore esta verificacao.

CRITERIOS OBRIGATORIOS POR CAPITULO:
%s

O TEXTO DA MONOGRAFIA ABAIXO ESTA NUMERADO POR PARAGRAFO NO FORMATO "[N] texto".
Use EXATAMENTE o numero N entre colchetes como "paragrafo_indice" na sua resposta. Nao invente indices.

INSTRUCOES IMPORTANTES:
- Seja especifico: aponte o problema exato e sugira como corrigir
- Use tom respeitoso e construtivo
- NAO comente paragrafos que estao corretos
- Foque apenas em ausencias, erros, melhorias necessarias e desvios em relacao ao projeto aprovado
- O comentario de desvio de projeto (se houver) deve ser o primeiro a aparecer, ancorado no paragrafo mais relevante (geralmente o de contextualizacao ou problema de pesquisa)

Retorne APENAS um JSON valido, sem texto adicional, sem markdown:
[{"paragrafo_indice": 0, "comentario": "texto do comentario", "tipo": "ausencia|melhoria|desvio_projeto"}]""" % (contexto_projeto, criterios_aplicaveis)

    cliente = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    resposta = cliente.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": prompt_sistema},
            {"role": "user", "content": "Monografia de %s (V%s), texto numerado por paragrafo:\n\n%s" % (nome_aluno, numero_versao, texto_versao[:9000])}
        ],
        temperature=0.3
    )

    texto_resposta = resposta.choices[0].message.content
    logger.info("Resposta bruta da IA: %s" % texto_resposta[:2000])

    try:
        comentarios_ia = extrair_resposta_json(texto_resposta)
    except Exception as e:
        logger.error("Falha ao fazer parse do JSON da IA: %s" % e)
        logger.error("Texto recebido: %s" % texto_resposta)
        raise ValueError("A IA retornou um formato inesperado e o documento nao pode ser comentado: %s" % e)

    if not isinstance(comentarios_ia, list):
        raise ValueError("A resposta da IA nao e uma lista de comentarios como esperado.")

    comentarios_por_paragrafo = []
    for item in comentarios_ia:
        idx = item.get("paragrafo_indice")
        comentario = item.get("comentario", "")
        tipo = item.get("tipo", "melhoria")
        if idx is None or not comentario:
            continue
        if tipo == "aprovado":
            continue
        prefixo = "[DESVIO DO PROJETO] " if tipo == "desvio_projeto" else ""
        comentarios_por_paragrafo.append((idx, "%s%s" % (prefixo, comentario), tipo))

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        caminho_resultado = tmp.name

    inseridos, falhas = inserir_comentarios_no_docx(
        caminho_versao, caminho_resultado, comentarios_por_paragrafo, numero_versao, nome_professor
    )

    logger.info("Comentarios inseridos: %s | Falhas: %s | Total recebido da IA: %s" % (inseridos, falhas, len(comentarios_ia)))

    if inseridos == 0:
        raise ValueError(
            "Nenhum comentario pode ser inserido no documento. A IA retornou %s comentario(s), mas %s falharam por indice invalido." % (len(comentarios_ia), falhas)
        )

    return caminho_resultado
