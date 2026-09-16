import os
import re
import unicodedata
import json
import logging
import shutil
import tempfile
import zipfile
from lxml import etree
from datetime import datetime
import openai
from extrator import extrair_texto_docx, extrair_texto_docx_completo

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("rubric-ai-masters")

openai.api_key = os.environ.get("OPENAI_API_KEY")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NSMAP = {"w": W_NS}

MODELO = "gpt-4o"
TAMANHO_MINIMO_COMENTARIO = 40
MAX_COMENTARIOS_POR_BLOCO = 8
LIMIAR_REPETICAO = 0.32
MAX_POR_TEMA = 2

CRITERIOS = {
    "introducao": """
A Introducao deve conter OBRIGATORIAMENTE todos os elementos abaixo. Aponte como erro/melhoria qualquer ausencia:

1. CONTEXTUALIZACAO DO TEMA: O aluno apresenta o tema escolhido com relevancia clara, preferencialmente com fundamentacao teorica (citacoes). Aponte se falta fundamentacao na contextualizacao.
2. PROBLEMA DE PESQUISA: Deve derivar da contextualizacao e OBRIGATORIAMENTE estar em forma de pergunta. Antes de apontar erro aqui, verifique se ha ponto de interrogacao no trecho. Se houver, o criterio esta atendido.
3. OBJETIVO GERAL: Deve derivar diretamente do problema de pesquisa.
4. OBJETIVOS ESPECIFICOS: Maximo de 3 ou 4. Antes de apontar excesso, CONTE os objetivos listados e cite a contagem no comentario.
5. JUSTIFICATIVA: Deve apresentar claramente por que a pesquisa e relevante. Aponte se estiver ausente ou vaga.
6. INDICACAO METODOLOGICA: Deve indicar brevemente qual sera a metodologia, sem detalhar.
7. PARAGRAFO DE SINTESE (ESTRUTURA): O ultimo paragrafo deve descrever como o trabalho esta organizado (ex: "O capitulo 2 aborda...").
8. FREQUENCIA DE CITACOES: O texto deve conter citacoes a cada 2 ou 3 paragrafos no minimo. Aponte trechos longos sem citacao.
""",
    "metodologia": """
A Metodologia deve conter OBRIGATORIAMENTE todos os elementos abaixo:

REGRAS GERAIS:
1. DESCRICAO DA METODOLOGIA: O tipo de pesquisa (bibliografica, de campo, etc.) deve estar claramente descrito.
2. FUNDAMENTACAO DA ESCOLHA: Minimo de 2 autores diferentes de metodologia cientifica. Antes de apontar falta, LISTE os autores de metodologia que encontrou no trecho.
3. PERIODO DE REALIZACAO: Indicar meses e ano cursados na disciplina de capstone.
4. PRINCIPIOS ETICOS: Paragrafo informando que o projeto foi aprovado pelo Comite de Etica institucional.

SE FOR PESQUISA BIBLIOGRAFICA (verificar todos os itens):
5. DESCRITORES/PALAVRAS-CHAVE usados na busca.
6. RECORTE TEMPORAL definido (ultimos 5 anos).
7. PLATAFORMAS DE BUSCA (ex: SciELO, Periodicos CAPES, BDTD, Google Academico).
8. IDIOMAS consultados (portugues, espanhol, ingles).
9. CRITERIOS DE INCLUSAO E EXCLUSAO das obras.
10. DESCRICAO QUANTITATIVA DO FUNIL: quantos materiais vieram da busca, quantos foram excluidos e quantos restaram.
11. QUADRO DE OBRAS RESULTANTES (pode estar aqui ou nos Resultados).
12. ESTRATEGIA DE ANALISE: leitura critica, sistematizacao e categorizacao.
13. LIMITACOES: indicar que e pesquisa teorica, nao empirica.

SE FOR PESQUISA DE CAMPO/EMPIRICA:
14. APROVACAO PREVIA no projeto de capstone e procedimentos de coleta e analise descritos.
""",
    "referencial_capitulo": """
Voce esta avaliando UM capitulo teorico. Verifique OBRIGATORIAMENTE, dentro deste capitulo:

1. CITACAO DIRETA: ha pelo menos uma citacao direta, com aspas e indicacao de pagina? Aponte se nao houver.
2. CITACAO INDIRETA: ha parafrase com credito ao autor, no formato (Autor, ano)? Aponte se o capitulo so reproduzir ideias sem credito.
3. FUNDAMENTACAO DAS AFIRMACOES: toda afirmacao relevante tem apoio teorico. Aponte trechos longos de opiniao sem respaldo.
4. FONTES RECENTES: predominam obras dos ultimos 5 anos, salvo classicos reconhecidos. Aponte se a bibliografia do capitulo for majoritariamente antiga.
5. CONFIABILIDADE DA FONTE: as fontes sao academicas (periodicos, livros, teses). Aponte uso de blog, site generico ou fonte sem autoria.
6. DIALOGO ENTRE AUTORES: os autores conversam entre si, com concordancias e divergencias. Aponte como erro a sequencia de citacoes isoladas, um autor por paragrafo, sem conexao.
7. VOZ DO ALUNO E POSICIONAMENTO CRITICO: ha costura autoral entre os conceitos, explicando como a teoria se aplica ao estudo. Aponte capitulo que seja apenas colagem de resumos.
8. ALINHAMENTO COM OS OBJETIVOS: o conteudo contribui para responder ao problema de pesquisa.
9. FOCO E DELIMITACAO: aponte conteudo periferico que nao serve aos objetivos.
10. COESAO, COERENCIA E ORTOGRAFIA: aponte problemas de encadeamento entre paragrafos e erros de escrita.
11. NORMA APA nas chamadas de autor dentro do texto.

REGRA DE AGRUPAMENTO, OBRIGATORIA:
Nao repita a mesma observacao em varios paragrafos. Se um problema se repete ao longo do capitulo (por
exemplo, varias citacoes antigas, ou varios trechos sem apoio teorico), gere UM UNICO comentario que trate
do problema no capitulo inteiro, citando dois ou tres exemplos dentro dele, ancorado no primeiro paragrafo
onde o problema aparece. Gere no maximo 6 comentarios neste capitulo, cada um sobre um criterio diferente.
""",
    "referencial_secao": """
Voce esta avaliando o CONJUNTO dos capitulos teoricos, nao o conteudo de um capitulo isolado.
Voce recebera a lista dos capitulos, com titulo e tamanho de cada um.

Verifique OBRIGATORIAMENTE:

1. LOGICA DA DIVISAO: a progressao entre os capitulos tem fio condutor claro (do geral ao especifico, do historico ao atual, ou outro criterio visivel). Aponte divisao arbitraria.
2. TRANSICAO ENTRE CAPITULOS: a sequencia dos titulos sugere encadeamento, nao capitulos soltos.
3. EQUILIBRIO DE EXTENSAO: aponte desproporcao grande entre capitulos, por exemplo um com o triplo do tamanho de outro.
4. QUANTIDADE: o referencial deve ter entre 2 e 5 capitulos. Aponte se fugir disso.
5. COBERTURA TEORICA: os titulos cobrem os temas exigidos pelos objetivos da pesquisa.

Gere no maximo 3 comentarios, ancorados no paragrafo do titulo do capitulo a que se referem.
""",
    "resultados": """
O capitulo de RESULTADOS E DISCUSSAO apresenta o que foi encontrado no levantamento bibliografico e aproxima
os autores em torno do que demonstram ter em comum.

Verifique OBRIGATORIAMENTE:

1. APRESENTACAO DOS RESULTADOS: achados apresentados de forma clara.
2. APROXIMACAO ENTRE AUTORES: afinidades e distanciamentos, no que concordam e no que discordam. Aponte como erro grave a simples listagem de autores.
3. ANALISE E INTERPRETACAO: significado dos achados, o "o que" e o "por que".
4. CONEXAO COM OBJETIVOS E COM O PROBLEMA de pesquisa.
5. SENTIDO DOS DADOS: numeros e quadros devem ganhar leitura, nao apenas ser exibidos.
6. QUADRO DE OBRAS RESULTANTES: se nao foi apresentado na Metodologia, deve estar aqui.
7. FREQUENCIA DE CITACOES: citacoes a cada 2 ou 3 paragrafos no minimo.
""",
    "conclusao": """
As CONSIDERACOES FINAIS devem ser claras, breves, objetivas e baseadas nos achados, e NAO devem conter
informacoes novas.

Verifique OBRIGATORIAMENTE:

1. RETOMADA DO TEMA, DO PROBLEMA E DOS OBJETIVOS de pesquisa.
2. ALCANCE DE CADA OBJETIVO: evidenciar se cada objetivo especifico foi atingido.
3. RESULTADOS OBTIDOS: destacar e discutir os achados do estudo.
4. IMPLICACOES dos resultados e contribuicao para o campo de estudo.
5. LIMITACOES da pesquisa e aspectos que podem ter influenciado os resultados.
6. CAUTELA NA GENERALIZACAO dos achados.
7. SUGESTOES DE PESQUISAS FUTURAS a partir das lacunas identificadas.
8. AUSENCIA DE CITACOES: esta secao nao deve conter citacoes de autores. Aponte como erro qualquer citacao.
9. NENHUMA INFORMACAO NOVA que nao tenha sido desenvolvida no corpo do trabalho.
""",
}

CRITERIO_ADERENCIA = """
Voce esta fazendo UMA unica verificacao: a aderencia da monografia ao Projeto de Capstone aprovado.

TRABALHE NESTA ORDEM, obrigatoriamente:

PASSO 1. Leia o PROJETO APROVADO e extraia, campo a campo:
   - TEMA aprovado, com todos os recortes que ele contem (area, disciplina, publico, etapa de ensino, contexto)
   - PROBLEMA DE PESQUISA aprovado, na integra
   - OBJETIVO GERAL aprovado
   - OBJETIVOS ESPECIFICOS aprovados
   - METODOLOGIA aprovada

PASSO 2. Leia o TITULO, o RESUMO e o trecho da MONOGRAFIA e extraia os mesmos cinco campos.

PASSO 3. Compare CAMPO A CAMPO, um de cada vez. Para cada campo, pergunte:
   - Todos os termos e recortes que existem no projeto continuam presentes na monografia?
   - REGRA CENTRAL: um recorte que existe no projeto e NAO aparece na monografia e DESVIO GRAVE, mesmo que o
     resto do texto esteja coerente. Exemplos de forma: se o projeto delimita uma disciplina, uma etapa de
     ensino, uma faixa etaria, uma regiao ou um publico especifico, e a monografia trata do tema sem essa
     delimitacao, isso e desvio e deve ser apontado.
   - A ausencia de um recorte e tao grave quanto a troca de tema. Nao conclua que houve apenas "ampliacao do
     escopo": trate como desvio.

PASSO 4. Para cada desvio, gere um comentario que cite TEXTUALMENTE, entre aspas, o que foi aprovado no
projeto e o que esta escrito na monografia, nesta ordem. Use o tipo "desvio_projeto".

Se, e somente se, os cinco campos corresponderem integralmente, retorne uma lista vazia.
Gere no maximo 4 comentarios.
"""

# Nomes que a tela envia e que apontam para um bloco de criterios.
ALIASES_CAPITULOS = {
    "discussao": "resultados",
    "resultados e discussao": "resultados",
    "fundamentacao": "referencial",
    "fundamentacao teorica": "referencial",
    "referencial teorico": "referencial",
    "consideracoes": "conclusao",
    "consideracoes finais": "conclusao",
    "conclusao": "conclusao",
    "introducao": "introducao",
    "metodologia": "metodologia",
}

NOMES_SECAO = {
    "introducao": ["introdução", "introducao"],
    "metodologia": ["metodologia", "materiais e métodos", "método", "metodo"],
    "resultados": ["resultados e discussão", "resultados e discussao", "resultados",
                   "discussão", "discussao"],
    "conclusao": ["considerações finais", "consideracoes finais", "conclusão", "conclusao"],
}
FIM_DO_CORPO = ["referências", "referencias", "glossário", "glossario",
                "apêndice", "apendice", "anexo", "anexos"]
MARCAS_PLACEHOLDER = ["lorem ipsum", "xxxxxxxx", "título do capítulo", "titulo do capitulo",
                      "título do capitulo", "xxxxxxxxxx"]


# ---------------------------------------------------------------- utilidades

def _sem_acento(t):
    return "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c))


def _normalizar(t):
    return t.strip().lower().rstrip(".:").strip()


def _chave_capitulo(nome):
    """Converte o nome que a tela envia (com acento) na chave interna de criterios."""
    n = _sem_acento(_normalizar(nome))
    return ALIASES_CAPITULOS.get(n, n)


def _sem_numeracao(texto):
    return re.sub(r"^\d+(\.\d+)*\s*\.?\s*", "", texto.strip())


def _classificar_nome(texto):
    n = _normalizar(_sem_numeracao(texto))
    for chave, nomes in NOMES_SECAO.items():
        for nome in nomes:
            if n == nome or n.startswith(nome + " ") or n.startswith(nome + "/"):
                return chave
    return None


def _e_titulo_numerado(texto):
    if len(texto) > 150:
        return None, None
    m = re.match(r"^(\d+)((?:\.\d+)*)\s*\.?\s*(\S.*)$", texto)
    if not m:
        return None, None
    resto = m.group(3).strip()
    if not resto or len(resto) > 140:
        return None, None
    profundidade = 1 + (m.group(2).count(".") if m.group(2) else 0)
    return profundidade, resto


def _linhas_numeradas(texto_numerado):
    saida = []
    for linha in texto_numerado.split("\n"):
        m = re.match(r"^\[(\d+)\]\s?(.*)$", linha)
        if m:
            saida.append((int(m.group(1)), m.group(2).strip()))
    return saida


def _pre_texto(texto_numerado, limite=40):
    """Devolve as linhas antes do Sumario: capa, titulo e resumo."""
    linhas = _linhas_numeradas(texto_numerado)
    corte = len(linhas)
    for pos, (_, texto) in enumerate(linhas):
        if _normalizar(texto) in ("sumário", "sumario"):
            corte = pos
            break
    return "\n".join("[%d] %s" % (idx, t) for idx, t in linhas[:corte][:limite])


FAMILIAS_TEMA = {
    "atualidade": ["antig", "recent", "atuali", "contemporane", "classic", "desatualiz"],
    "citacao_direta": ["citacao direta", "aspas", "indicacao de pagina"],
    "fundamentacao": ["sem apoio", "fundamenta", "respaldo", "sem citacao", "sem fonte",
                      "nao cita", "sem referencia"],
    "dialogo": ["dialogo", "conversam", "isolad", "confronto", "concordam", "divergem"],
    "voz_autoral": ["voz do aluno", "posicionamento critico", "autoral", "colagem", "voz autoral"],
    "norma_apa": ["norma apa", "formatacao apa", "padrao apa"],
    "coesao": ["coesao", "coerencia", "ortograf", "encadeamento", "transicao"],
    "confiabilidade": ["blog", "sem autoria", "site generico", "fonte nao academica"],
}


def _familia(texto):
    t = _sem_acento(texto.lower())
    for nome, marcas in FAMILIAS_TEMA.items():
        if any(m in t for m in marcas):
            return nome
    return None


def _assinatura(texto):
    palavras = re.findall(r"[a-z]{5,}", _sem_acento(texto.lower()))
    return set(palavras)


def _limitar_repeticao(validos, rotulo):
    """Descarta comentario muito parecido com outro ja aceito no mesmo bloco."""
    aceitos = []
    assinaturas = []
    por_tema = {}
    for idx, texto, tipo in validos:
        if len(aceitos) >= MAX_COMENTARIOS_POR_BLOCO:
            logger.info("[%s] teto de %s comentarios atingido, restante descartado"
                        % (rotulo, MAX_COMENTARIOS_POR_BLOCO))
            break
        tema = _familia(texto)
        if tema and por_tema.get(tema, 0) >= MAX_POR_TEMA:
            logger.info("[%s] tema %s ja tem %s comentarios, descartado: %r"
                        % (rotulo, tema, MAX_POR_TEMA, texto[:60]))
            continue
        atual = _assinatura(texto)
        repetido = False
        for anterior in assinaturas:
            uniao = atual | anterior
            if uniao and len(atual & anterior) / float(len(uniao)) >= LIMIAR_REPETICAO:
                repetido = True
                break
        if repetido:
            logger.info("[%s] comentario repetido descartado: %r" % (rotulo, texto[:60]))
            continue
        assinaturas.append(atual)
        if tema:
            por_tema[tema] = por_tema.get(tema, 0) + 1
        aceitos.append((idx, texto, tipo))
    return aceitos


def tem_placeholder(texto):
    t = texto.lower()
    return any(m in t for m in MARCAS_PLACEHOLDER)


def fatiar_documento(texto_numerado):
    """Divide o texto numerado em blocos por capitulo, preservando os indices de paragrafo."""
    linhas = _linhas_numeradas(texto_numerado)
    if not linhas:
        return []

    inicio = 0
    for pos, (_, texto) in enumerate(linhas):
        if _normalizar(texto) in ("sumário", "sumario"):
            inicio = pos + 1
    if inicio == 0:
        for pos, (_, texto) in enumerate(linhas):
            if _classificar_nome(texto) == "introducao":
                inicio = pos
                break

    fim = len(linhas)
    for pos in range(inicio, len(linhas)):
        if _normalizar(_sem_numeracao(linhas[pos][1])) in FIM_DO_CORPO:
            fim = pos
            break

    corpo = linhas[inicio:fim]
    if not corpo:
        return []

    marcos = []
    capitulo_atual = None
    for pos, (_, texto) in enumerate(corpo):
        chave = _classificar_nome(texto)
        nivel, resto = _e_titulo_numerado(texto)
        numero = re.match(r"^(\d+)", texto.strip()).group(1) if nivel else None
        if chave and (nivel is None or nivel == 1):
            marcos.append((pos, chave, texto))
            capitulo_atual = numero
        elif nivel == 1:
            marcos.append((pos, _classificar_nome(resto) or "referencial", texto))
            capitulo_atual = numero
        elif nivel and nivel >= 2 and numero != capitulo_atual:
            marcos.append((pos, "referencial", texto))
            capitulo_atual = numero

    if not marcos:
        return []

    blocos = []
    for i, (pos, chave, titulo) in enumerate(marcos):
        pos_fim = marcos[i + 1][0] if i + 1 < len(marcos) else len(corpo)
        trecho = corpo[pos:pos_fim]
        blocos.append({
            "chave": chave,
            "titulo": titulo,
            "idx_inicio": trecho[0][0],
            "idx_fim": trecho[-1][0],
            "paragrafos": len(trecho),
            "palavras": sum(len(t.split()) for _, t in trecho),
            "texto": "\n".join("[%d] %s" % (idx, t) for idx, t in trecho),
        })
    return blocos


def extrair_resposta_json(texto_resposta):
    """Extrai e faz parse do JSON retornado pela IA, tolerando blocos markdown."""
    texto = (texto_resposta or "").strip()
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


# ---------------------------------------------------------- chamada ao modelo

INSTRUCOES_COMUNS = """
O TEXTO ABAIXO ESTA NUMERADO POR PARAGRAFO NO FORMATO "[N] texto".
Use EXATAMENTE o numero N entre colchetes como "paragrafo_indice". Nao invente indices e nao use indices
que nao aparecam no texto recebido.

REGRAS DE ESCRITA DOS COMENTARIOS:
- Escreva em portugues brasileiro, com tom respeitoso e construtivo, como um orientador falando com o aluno.
- Cada comentario deve ter no minimo duas frases: o que esta faltando ou errado, e como corrigir.
- Cada comentario deve CITAR entre aspas um trecho curto do texto do aluno que motivou a observacao, ou
  dizer explicitamente que o elemento nao foi encontrado no capitulo.
- NAO afirme que algo esta ausente sem ter procurado no texto recebido.
- NAO comente paragrafos corretos. Foque em ausencias, erros e melhorias.
- NUNCA escreva apenas o tipo do comentario como texto. O campo "comentario" e o texto que o aluno vai ler.

Retorne APENAS um JSON valido, sem markdown e sem texto adicional:
[{"paragrafo_indice": 0, "comentario": "texto do comentario", "tipo": "melhoria"}]
Se nao houver nada a apontar, retorne [].
"""


def _chamar_ia(cliente, instrucao_criterios, texto_usuario, rotulo):
    hoje = datetime.now()
    prompt_sistema = (
        "Voce e avaliador especialista de monografias de mestrado da Must University.\n"
        "Norma academica: APA.\n"
        "A data de hoje e %s. O ano corrente e %d. Considere \"ultimos 5 anos\" como %d a %d. "
        "NUNCA afirme que um ano igual ou anterior a %d ainda nao ocorreu.\n\n"
        % (hoje.strftime("%d/%m/%Y"), hoje.year, hoje.year - 5, hoje.year, hoje.year)
        + instrucao_criterios
        + "\n"
        + INSTRUCOES_COMUNS
    )
    try:
        resposta = cliente.chat.completions.create(
            model=MODELO,
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": texto_usuario[:60000]},
            ],
            temperature=0.2,
        )
        bruto = resposta.choices[0].message.content
        itens = extrair_resposta_json(bruto)
        if not isinstance(itens, list):
            logger.error("[%s] resposta nao e lista: %s" % (rotulo, bruto[:500]))
            return []
        return itens
    except Exception as e:
        logger.error("[%s] falha na chamada ou no parse: %s" % (rotulo, e))
        return []


def _validar(itens, rotulo, faixa=None):
    """Descarta comentario quebrado, vazio, com indice invalido ou igual ao tipo."""
    tipos = {"ausencia", "melhoria", "desvio_projeto", "aprovado", "erro", "placeholder"}
    validos = []
    for item in itens:
        if not isinstance(item, dict):
            continue
        idx = item.get("paragrafo_indice")
        comentario = (item.get("comentario") or "").strip()
        tipo = (item.get("tipo") or "melhoria").strip()
        if not isinstance(idx, int):
            continue
        if tipo == "aprovado":
            continue
        if comentario.lower() in tipos or len(comentario) < TAMANHO_MINIMO_COMENTARIO:
            logger.warning("[%s] comentario descartado por texto invalido: %r" % (rotulo, comentario))
            continue
        if faixa and not (faixa[0] <= idx <= faixa[1]):
            logger.warning("[%s] indice %s fora da faixa %s, descartado" % (rotulo, idx, faixa))
            continue
        validos.append((idx, comentario, tipo))
    return validos


# ------------------------------------------------- insercao dos comentarios

def inserir_comentarios_no_docx(caminho_entrada, caminho_saida, comentarios_por_paragrafo,
                                numero_versao, nome_professor="Professor(a)"):
    """Insere comentarios nativos do Word manipulando diretamente o pacote ZIP do .docx."""
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
                logger.warning("Indice fora do intervalo: %s (total: %s)" % (idx, len(paragrafos_xml)))
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
        etree.ElementTree(comments_root).write(comments_xml_path, xml_declaration=True,
                                               encoding="UTF-8", standalone=True)
        tree.write(document_xml_path, xml_declaration=True, encoding="UTF-8", standalone=True)

        _garantir_relacionamento_comentarios(pasta_temp)
        _garantir_content_type_comentarios(pasta_temp)

        if os.path.exists(caminho_saida):
            os.remove(caminho_saida)
        with zipfile.ZipFile(caminho_saida, 'w', zipfile.ZIP_DEFLATED) as zf:
            for raiz, _, arquivos in os.walk(pasta_temp):
                for nome_arquivo in arquivos:
                    caminho_completo = os.path.join(raiz, nome_arquivo)
                    zf.write(caminho_completo, os.path.relpath(caminho_completo, pasta_temp))

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
        numeros = [int(i.replace("rId", "")) for i in ids
                   if i.startswith("rId") and i.replace("rId", "").isdigit()]
        rel = etree.SubElement(root, "{%s}Relationship" % RELS_NS)
        rel.set("Id", "rId%s" % (max(numeros, default=0) + 1))
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
        rel.set("Target", "comments.xml")

    tree.write(rels_path, xml_declaration=True, encoding="UTF-8", standalone=True)


def _garantir_content_type_comentarios(pasta_temp):
    ct_path = os.path.join(pasta_temp, "[Content_Types].xml")
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(ct_path, parser)
    root = tree.getroot()
    ja_existe = any(el.get("PartName") == "/word/comments.xml" for el in root
                    if el.tag == "{%s}Override" % CT_NS)
    if not ja_existe:
        override = etree.SubElement(root, "{%s}Override" % CT_NS)
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType",
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    tree.write(ct_path, xml_declaration=True, encoding="UTF-8", standalone=True)


# ------------------------------------------------------------ fluxo principal

async def processar_documento(caminho_versao, caminho_projeto, nome_aluno, numero_versao,
                              capitulos, nome_professor="Professor(a)"):
    texto_versao = extrair_texto_docx(caminho_versao)
    linhas = _linhas_numeradas(texto_versao)
    if not linhas:
        raise ValueError("Nao foi possivel ler o texto do documento enviado.")
    ultimo_indice = linhas[-1][0]

    # capitulos que a professora marcou na tela
    selecionados = []
    for cap in capitulos or []:
        chave = _chave_capitulo(cap)
        if chave in ("introducao", "metodologia", "referencial", "resultados", "conclusao") \
                and chave not in selecionados:
            selecionados.append(chave)
    if not selecionados:
        selecionados = ["introducao", "metodologia", "referencial"]

    blocos = fatiar_documento(texto_versao)
    logger.info("Capitulos marcados: %s | Blocos detectados: %s"
                % (selecionados, [(b["chave"], b["titulo"][:40], b["paragrafos"]) for b in blocos]))

    cliente = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    comentarios = []

    # 1. aderencia ao projeto de capstone
    if caminho_projeto:
        texto_projeto = extrair_texto_docx_completo(caminho_projeto)
        blocos_base = [b for b in blocos if b["chave"] in ("introducao", "metodologia")]
        trecho_base = "\n".join(b["texto"] for b in blocos_base) or texto_versao[:20000]
        capa = _pre_texto(texto_versao)
        usuario = ("PROJETO DE CAPSTONE APROVADO:\n%s\n\n"
                   "TITULO E RESUMO DA MONOGRAFIA:\n%s\n\n"
                   "INTRODUCAO E METODOLOGIA DA MONOGRAFIA DE %s (V%s):\n%s"
                   % (texto_projeto[:20000], capa, nome_aluno, numero_versao, trecho_base))
        itens = _chamar_ia(cliente, CRITERIO_ADERENCIA, usuario, "aderencia")
        validos = _validar(itens, "aderencia")
        for idx, texto, _tipo in validos:
            comentarios.append((idx, "[DESVIO DO PROJETO] " + texto, "desvio_projeto"))
        logger.info("[aderencia] %s comentario(s)" % len(validos))

    # 2. uma chamada por capitulo marcado
    for chave in selecionados:
        blocos_da_chave = [b for b in blocos if b["chave"] == chave]

        if not blocos_da_chave:
            comentarios.append((
                ultimo_indice,
                "O capitulo de %s foi marcado como presente nesta versao, mas nao foi localizado no "
                "documento. Verifique se o titulo do capitulo esta escrito no texto e se o conteudo "
                "foi enviado." % chave.capitalize(),
                "ausencia"))
            logger.info("[%s] capitulo marcado e nao encontrado" % chave)
            continue

        criterio = CRITERIOS["referencial_capitulo"] if chave == "referencial" else CRITERIOS[chave]

        for bloco in blocos_da_chave:
            rotulo = "%s/%s" % (chave, bloco["titulo"][:30])

            if tem_placeholder(bloco["texto"]):
                comentarios.append((
                    bloco["idx_inicio"],
                    "Este capitulo ainda esta com o texto padrao do modelo, com marcacoes como "
                    "\"lorem ipsum\" ou \"Xxxxxxxx\". Substitua o conteudo do modelo pelo texto da "
                    "pesquisa antes de submeter a proxima versao.",
                    "placeholder"))
                logger.info("[%s] placeholder detectado, capitulo nao avaliado" % rotulo)
                continue

            usuario = ("Capitulo: %s\nMonografia de %s (V%s)\n\n%s"
                       % (bloco["titulo"], nome_aluno, numero_versao, bloco["texto"]))
            itens = _chamar_ia(cliente, criterio, usuario, rotulo)
            validos = _limitar_repeticao(
                _validar(itens, rotulo, faixa=(bloco["idx_inicio"], bloco["idx_fim"])), rotulo)
            comentarios.extend(validos)
            logger.info("[%s] %s comentario(s) em %s paragrafos"
                        % (rotulo, len(validos), bloco["paragrafos"]))

    # 3. visao do conjunto dos capitulos teoricos
    teoricos = [b for b in blocos if b["chave"] == "referencial" and not tem_placeholder(b["texto"])]
    if "referencial" in selecionados and len(teoricos) >= 2:
        resumo = "\n".join("[%d] %s (%d paragrafos, %d palavras)"
                           % (b["idx_inicio"], b["titulo"], b["paragrafos"], b["palavras"])
                           for b in teoricos)
        usuario = ("Capitulos teoricos da monografia de %s (V%s):\n\n%s" % (nome_aluno, numero_versao, resumo))
        itens = _chamar_ia(cliente, CRITERIOS["referencial_secao"], usuario, "referencial/conjunto")
        validos = _validar(itens, "referencial/conjunto",
                           faixa=(teoricos[0]["idx_inicio"], teoricos[-1]["idx_fim"]))
        comentarios.extend(validos)
        logger.info("[referencial/conjunto] %s comentario(s)" % len(validos))

    # 4. remove duplicados e ordena pela posicao no documento
    vistos = set()
    finais = []
    for idx, texto, tipo in comentarios:
        chave_dedup = (idx, texto[:60].lower())
        if chave_dedup in vistos:
            continue
        vistos.add(chave_dedup)
        finais.append((idx, texto, tipo))
    finais.sort(key=lambda c: c[0])

    logger.info("TOTAL de comentarios validos: %s" % len(finais))

    if not finais:
        raise ValueError(
            "A analise nao gerou nenhum comentario valido para este documento. "
            "Verifique se o arquivo enviado e a dissertacao e se os capitulos marcados existem no texto."
        )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        caminho_resultado = tmp.name

    inseridos, falhas = inserir_comentarios_no_docx(
        caminho_versao, caminho_resultado, finais, numero_versao, nome_professor
    )
    logger.info("Comentarios inseridos: %s | Falhas: %s | Validos: %s"
                % (inseridos, falhas, len(finais)))

    if inseridos == 0:
        raise ValueError(
            "Nenhum comentario pode ser inserido no documento. Foram gerados %s comentario(s) validos, "
            "mas %s falharam por indice invalido." % (len(finais), falhas)
        )

    return caminho_resultado