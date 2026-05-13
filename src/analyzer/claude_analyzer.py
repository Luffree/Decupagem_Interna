"""
Claude API analyzer.
Sends extracted document content to Claude and receives structured decupagem data.
"""
import json
import anthropic

SYSTEM_PROMPT = """Você é um agente especialista em decupagem técnica de projetos para orçamento, produção e conferência.

Seu trabalho é analisar o conteúdo bruto extraído de arquivos PDF ou PPT e transformá-lo em uma estrutura de planilha técnica detalhada.

Seu foco é identificar, separar, estruturar e organizar cada item do projeto linha por linha, sem consolidar itens diferentes na mesma célula ou na mesma linha. Você deve agir como um analista técnico de cenografia, produção, orçamento e detalhamento executivo.

## Regras críticas de extração

- Preserve a divisão por página/slide de origem.
- Separe cada objeto, peça, comunicação visual, mobiliário, acabamento, material, arte, estrutura, compra, locação ou pendência em sua própria linha.
- Nunca junte itens diferentes em uma única linha.
- Quando houver item principal e subcomponentes, crie uma linha para o item principal e linhas separadas para cada subcomponente.
- Quando uma informação não estiver clara, use "Não informado" ou "A confirmar".
- Não invente medidas, quantidades, materiais, fornecedores ou status.
- Preserve a unidade original exatamente como aparece.
- Mantenha a linguagem técnica do projeto o mais fiel possível.

## Formato de saída OBRIGATÓRIO

Responda APENAS com um JSON válido seguindo exatamente esta estrutura:

{
  "resumo": {
    "cliente": "string",
    "projeto": "string",
    "total_paginas_slides": number,
    "total_itens": number,
    "secoes": [{"nome": "string", "quantidade_itens": number}],
    "fornecedores": [{"nome": "string", "quantidade_itens": number}],
    "status_summary": [{"status": "string", "quantidade": number}],
    "observacoes_gerais": "string"
  },
  "itens_detalhados": [
    {
      "id": "string",
      "pagina_slide_origem": "string",
      "secao": "string",
      "sub_secao_ambiente": "string",
      "fornecedor": "string",
      "categoria": "string",
      "item": "string",
      "descricao_detalhada": "string",
      "quantidade": "string",
      "unidade": "string",
      "medida_original": "string",
      "largura": "string",
      "altura": "string",
      "profundidade": "string",
      "comprimento": "string",
      "espessura": "string",
      "area": "string",
      "metro_linear": "string",
      "material": "string",
      "acabamento_cor": "string",
      "tipo_producao": "string",
      "status_arte": "string",
      "status_compra_locacao": "string",
      "responsavel_area": "string",
      "observacoes_tecnicas": "string",
      "pendencia_acao_necessaria": "string",
      "nivel_atencao": "Baixo|Médio|Alto",
      "fonte_informacao": "string"
    }
  ],
  "por_secao": [
    {
      "secao": "string",
      "descricao_secao": "string",
      "ids_itens": ["string"],
      "paginas_slides": ["string"],
      "observacoes": "string"
    }
  ],
  "medicoes": [
    {
      "item": "string",
      "id_item": "string",
      "pagina_slide": "string",
      "medida_original": "string",
      "unidade_original": "string",
      "medida_metros": "string",
      "tipo_medida": "string",
      "confiabilidade": "Alta|Média|Baixa",
      "observacao_confiabilidade": "string"
    }
  ],
  "pendencias": [
    {
      "id_item": "string",
      "item": "string",
      "pagina_slide": "string",
      "secao": "string",
      "descricao_pendencia": "string",
      "tipo_pendencia": "string",
      "nivel_urgencia": "Baixo|Médio|Alto|Crítico",
      "responsavel": "string",
      "prazo": "string"
    }
  ],
  "legenda_status": [
    {
      "cor_hex": "string",
      "significado": "string",
      "fornecedor": "string",
      "status_relacionado": "string",
      "observacoes": "string"
    }
  ]
}

Preencha TODOS os campos. Se não houver informação, use "Não informado" ou "A confirmar".
Para medidas em metros, converta com precisão quando possível; caso contrário, use "A confirmar".
NÃO inclua nenhum texto fora do JSON. A resposta deve começar com { e terminar com }."""


def _build_user_message(doc_dict: dict, doc_type: str) -> str:
    total_key = "total_pages" if doc_type == "pdf" else "total_slides"
    items_key = "pages" if doc_type == "pdf" else "slides"
    item_key = "page_number" if doc_type == "pdf" else "slide_number"

    total = doc_dict.get(total_key, 0)
    items = doc_dict.get(items_key, [])

    lines = [
        f"Tipo de arquivo: {doc_type.upper()}",
        f"Nome do arquivo: {doc_dict.get('filename', 'Não informado')}",
        f"Total de {'páginas' if doc_type == 'pdf' else 'slides'}: {total}",
        "",
    ]

    if doc_dict.get("metadata"):
        lines.append("=== METADADOS DO ARQUIVO ===")
        for k, v in doc_dict["metadata"].items():
            if v:
                lines.append(f"{k}: {v}")
        lines.append("")

    for item in items:
        num = item.get(item_key, "?")
        label = "PÁGINA" if doc_type == "pdf" else "SLIDE"
        title = item.get("title", "")

        lines.append(f"{'='*60}")
        lines.append(f"=== {label} {num}" + (f" — {title}" if title else "") + " ===")
        lines.append(f"{'='*60}")

        text = item.get("text", "")
        if text and text.strip():
            lines.append("[TEXTO PRINCIPAL]")
            lines.append(text.strip())
            lines.append("")

        text_blocks = item.get("text_blocks", [])
        if text_blocks:
            lines.append("[BLOCOS DE TEXTO]")
            for b in text_blocks:
                if b.strip():
                    lines.append(f"  • {b.strip()}")
            lines.append("")

        notes = item.get("notes", "")
        if notes and notes.strip():
            lines.append("[NOTAS DO SLIDE]")
            lines.append(notes.strip())
            lines.append("")

        tables = item.get("tables", [])
        if tables:
            lines.append(f"[TABELAS: {len(tables)} encontrada(s)]")
            for t_idx, table in enumerate(tables):
                lines.append(f"  Tabela {t_idx + 1}:")
                for row in table:
                    lines.append("    | " + " | ".join(str(c) for c in row) + " |")
            lines.append("")

        annotations = item.get("annotations", [])
        if annotations:
            lines.append("[ANOTAÇÕES / CHAMADAS]")
            for ann in annotations:
                if ann.get("content"):
                    lines.append(f"  [{ann.get('type', '?')}] {ann['content']}")
            lines.append("")

        images = item.get("image_descriptions", [])
        if images:
            lines.append("[IMAGENS NA PÁGINA]")
            for img in images:
                lines.append(f"  • {img}")
            lines.append("")

        shapes = item.get("shapes_summary", [])
        if shapes:
            lines.append("[FORMAS COM TEXTO/COR]")
            for sh in shapes:
                parts = []
                if sh.get("name"):
                    parts.append(f"Nome: {sh['name']}")
                if sh.get("text"):
                    parts.append(f"Texto: {sh['text']}")
                if sh.get("fill_color"):
                    parts.append(f"Cor preenchimento: #{sh['fill_color']}")
                if parts:
                    lines.append("  • " + " | ".join(parts))
            lines.append("")

        placeholders = item.get("placeholder_data", [])
        if placeholders:
            lines.append("[PLACEHOLDERS]")
            for ph in placeholders:
                if ph.get("text"):
                    lines.append(f"  [{ph.get('type', '?')}] {ph['text']}")
            lines.append("")

        hyperlinks = item.get("hyperlinks", [])
        if hyperlinks:
            lines.append("[HYPERLINKS]")
            for link in hyperlinks:
                lines.append(f"  • {link}")
            lines.append("")

    lines.append("")
    lines.append("Analise TODO o conteúdo acima e gere a decupagem técnica completa no formato JSON especificado.")

    return "\n".join(lines)


def analyze_document(
    doc_dict: dict,
    doc_type: str,
    api_key: str,
    progress_callback=None,
) -> dict:
    """
    Send document content to Claude and return structured decupagem data.

    doc_type: "pdf" or "pptx"
    progress_callback: optional callable(message: str)
    """
    client = anthropic.Anthropic(api_key=api_key)

    user_message = _build_user_message(doc_dict, doc_type)

    if progress_callback:
        progress_callback("Enviando conteúdo para análise pelo Claude...")

    # Use extended thinking for deep analysis + prompt caching for system prompt
    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=16000,
        system=[
            {
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[{"role": "user", "content": user_message}],
    )

    if progress_callback:
        progress_callback("Análise concluída. Processando resposta...")

    raw_text = message.content[0].text if message.content else "{}"

    # Find the JSON boundaries robustly
    start = raw_text.find("{")
    end = raw_text.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError("Claude não retornou JSON válido na resposta.")

    json_str = raw_text[start:end]

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise ValueError(f"Erro ao interpretar JSON da resposta: {e}\n\nResposta recebida:\n{json_str[:500]}")

    return data
