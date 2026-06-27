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
        contexto_projeto = f"PROJETO DE CAPSTONE (documento de referência aprovado pelo professor):\n{texto_projeto[:3000]}"
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

CRITÉRIOS OBRIGATÓRIOS POR CAPÍTULO:
{criterios_aplicaveis}

INSTRUÇÕES IMPORTANTES:
- Seja específico: aponte o problema exato e sugira como corrigir
- Use tom respeitoso e construtivo
- NÃO comente parágrafos que estão corretos (tipo "aprovado")
- Foque apenas em ausências, erros e melhorias necessárias
- Considere o contexto do projeto de capstone ao avaliar

Retorne APENAS um JSON válido, sem texto adicional, sem markdown:
[{{"paragrafo_indice": 0, "comentario": "texto do comentário", "tipo": "ausencia|melhoria"}}]"""

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




    
