"""
Excel generator for decupagem output.
Creates a formatted .xlsx with 6 sheets: Resumo, Itens Detalhados,
Por Seção, Medições, Pendências, Legenda/Status.
"""
import io
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.styles.numbers import FORMAT_TEXT
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.formatting.rule import (
    CellIsRule, FormulaRule, ColorScaleRule, DataBarRule
)


# ── Palette ──────────────────────────────────────────────────────────────────
COL_HEADER_BG    = "1F3864"   # dark navy  – headers
COL_HEADER_FG    = "FFFFFF"
COL_SECTION_BG   = "2E75B6"   # blue       – section rows
COL_SECTION_FG   = "FFFFFF"
COL_ALT_ROW      = "DCE6F1"   # light blue – alternate rows
COL_ALTO         = "FF6B6B"   # red-ish    – nivel_atencao Alto
COL_MEDIO        = "FFD93D"   # yellow     – nivel_atencao Médio
COL_BAIXO        = "6BCB77"   # green      – nivel_atencao Baixo
COL_PENDENCIA    = "FFA07A"   # salmon     – pendencias
COL_APROVADO     = "90EE90"   # light green– aprovado
COL_RESUMO_LABEL = "E8F0FE"
COL_BORDER       = "BDC3C7"


def _solid(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _border(style="thin"):
    s = Side(style=style, color=COL_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)


def _header_font(bold=True, size=10, color=COL_HEADER_FG):
    return Font(name="Calibri", bold=bold, size=size, color=color)


def _body_font(bold=False, size=9, color="000000"):
    return Font(name="Calibri", bold=bold, size=size, color=color)


def _wrap_align(horizontal="left", vertical="top"):
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=True)


def _apply_header_row(ws, headers: list[str], row=1,
                      bg=COL_HEADER_BG, fg=COL_HEADER_FG, height=22):
    ws.row_dimensions[row].height = height
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.font = Font(name="Calibri", bold=True, size=10, color=fg)
        cell.fill = _solid(bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _border()


def _set_col_widths(ws, widths: dict[int, float]):
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def _freeze(ws, cell="A2"):
    ws.freeze_panes = cell


def _autofilter(ws, first_row=1):
    ws.auto_filter.ref = ws.dimensions


def _add_table(ws, name: str, ref: str, style="TableStyleMedium9"):
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name=style,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def _color_row(ws, row: int, n_cols: int, hex_color: str):
    for col in range(1, n_cols + 1):
        ws.cell(row=row, column=col).fill = _solid(hex_color)


# ── Sheet: Resumo ─────────────────────────────────────────────────────────────
def _build_resumo(ws, resumo: dict, filename: str):
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:D1")
    cell = ws["A1"]
    cell.value = "DECUPAGEM TÉCNICA — RESUMO DO PROJETO"
    cell.font = Font(name="Calibri", bold=True, size=14, color=COL_HEADER_FG)
    cell.fill = _solid(COL_HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    pairs = [
        ("Arquivo analisado", filename),
        ("Cliente", resumo.get("cliente", "Não informado")),
        ("Projeto", resumo.get("projeto", "Não informado")),
        ("Total de páginas/slides", resumo.get("total_paginas_slides", "Não informado")),
        ("Total de itens extraídos", resumo.get("total_itens", "Não informado")),
        ("Observações gerais", resumo.get("observacoes_gerais", "")),
    ]

    row = 3
    for label, value in pairs:
        ws.cell(row=row, column=1, value=label).font = Font(name="Calibri", bold=True, size=10, color="1F3864")
        ws.cell(row=row, column=1).fill = _solid(COL_RESUMO_LABEL)
        ws.cell(row=row, column=1).border = _border()
        val_cell = ws.cell(row=row, column=2, value=str(value))
        val_cell.font = _body_font(size=10)
        val_cell.alignment = _wrap_align()
        val_cell.border = _border()
        ws.row_dimensions[row].height = 18
        row += 1

    # Seções
    row += 1
    ws.cell(row=row, column=1, value="ITENS POR SEÇÃO").font = Font(name="Calibri", bold=True, size=10, color=COL_HEADER_FG)
    ws.cell(row=row, column=1).fill = _solid(COL_SECTION_BG)
    ws.cell(row=row, column=2).fill = _solid(COL_SECTION_BG)
    ws.merge_cells(f"A{row}:B{row}")
    row += 1
    _apply_header_row(ws, ["Seção", "Qtd. Itens"], row=row, bg="2E75B6")
    row += 1
    for secao in resumo.get("secoes", []):
        ws.cell(row=row, column=1, value=secao.get("nome", "")).font = _body_font()
        ws.cell(row=row, column=1).border = _border()
        ws.cell(row=row, column=2, value=secao.get("quantidade_itens", 0)).font = _body_font()
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).border = _border()
        if row % 2 == 0:
            ws.cell(row=row, column=1).fill = _solid(COL_ALT_ROW)
            ws.cell(row=row, column=2).fill = _solid(COL_ALT_ROW)
        row += 1

    # Fornecedores
    row += 1
    ws.cell(row=row, column=1, value="ITENS POR FORNECEDOR").font = Font(name="Calibri", bold=True, size=10, color=COL_HEADER_FG)
    ws.cell(row=row, column=1).fill = _solid(COL_SECTION_BG)
    ws.cell(row=row, column=2).fill = _solid(COL_SECTION_BG)
    ws.merge_cells(f"A{row}:B{row}")
    row += 1
    _apply_header_row(ws, ["Fornecedor", "Qtd. Itens"], row=row, bg="2E75B6")
    row += 1
    for forn in resumo.get("fornecedores", []):
        ws.cell(row=row, column=1, value=forn.get("nome", "")).font = _body_font()
        ws.cell(row=row, column=1).border = _border()
        ws.cell(row=row, column=2, value=forn.get("quantidade_itens", 0)).font = _body_font()
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).border = _border()
        if row % 2 == 0:
            ws.cell(row=row, column=1).fill = _solid(COL_ALT_ROW)
            ws.cell(row=row, column=2).fill = _solid(COL_ALT_ROW)
        row += 1

    # Status
    row += 1
    ws.cell(row=row, column=1, value="STATUS DOS ITENS").font = Font(name="Calibri", bold=True, size=10, color=COL_HEADER_FG)
    ws.cell(row=row, column=1).fill = _solid(COL_SECTION_BG)
    ws.cell(row=row, column=2).fill = _solid(COL_SECTION_BG)
    ws.merge_cells(f"A{row}:B{row}")
    row += 1
    _apply_header_row(ws, ["Status", "Qtd."], row=row, bg="2E75B6")
    row += 1
    for st in resumo.get("status_summary", []):
        ws.cell(row=row, column=1, value=st.get("status", "")).font = _body_font()
        ws.cell(row=row, column=1).border = _border()
        ws.cell(row=row, column=2, value=st.get("quantidade", 0)).font = _body_font()
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).border = _border()
        row += 1

    _set_col_widths(ws, {1: 35, 2: 55, 3: 20, 4: 20})


# ── Sheet: Itens Detalhados ───────────────────────────────────────────────────
ITENS_HEADERS = [
    "ID", "Página/Slide", "Seção", "Sub-seção/Ambiente",
    "Fornecedor", "Categoria", "Item", "Descrição Detalhada",
    "Qtd.", "Unid.", "Medida Original",
    "Largura", "Altura", "Profundidade", "Comprimento", "Espessura",
    "Área", "Metro Linear",
    "Material", "Acabamento/Cor", "Tipo Produção",
    "Status Arte", "Status Compra/Locação",
    "Responsável/Área", "Observações Técnicas",
    "Pendência/Ação", "Nível Atenção", "Fonte Informação",
]

ITENS_FIELDS = [
    "id", "pagina_slide_origem", "secao", "sub_secao_ambiente",
    "fornecedor", "categoria", "item", "descricao_detalhada",
    "quantidade", "unidade", "medida_original",
    "largura", "altura", "profundidade", "comprimento", "espessura",
    "area", "metro_linear",
    "material", "acabamento_cor", "tipo_producao",
    "status_arte", "status_compra_locacao",
    "responsavel_area", "observacoes_tecnicas",
    "pendencia_acao_necessaria", "nivel_atencao", "fonte_informacao",
]

ITENS_WIDTHS = {
    1: 8, 2: 12, 3: 22, 4: 22,
    5: 18, 6: 16, 7: 28, 8: 45,
    9: 7, 10: 7, 11: 20,
    12: 9, 13: 9, 14: 12, 15: 13, 16: 11,
    17: 9, 18: 12,
    19: 20, 20: 20, 21: 16,
    22: 16, 23: 18,
    24: 18, 25: 40,
    26: 35, 27: 13, 28: 20,
}

NIVEL_COLORS = {
    "alto":  COL_ALTO,
    "médio": COL_MEDIO,
    "medio": COL_MEDIO,
    "baixo": COL_BAIXO,
}


def _build_itens(ws, itens: list[dict]):
    _apply_header_row(ws, ITENS_HEADERS, row=1, height=24)
    _freeze(ws, "A2")
    _set_col_widths(ws, ITENS_WIDTHS)

    for r_idx, item in enumerate(itens, start=2):
        nivel = str(item.get("nivel_atencao", "")).strip().lower()
        row_color = NIVEL_COLORS.get(nivel, None)

        for c_idx, field in enumerate(ITENS_FIELDS, start=1):
            value = item.get(field, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=str(value) if value else "")
            cell.font = _body_font()
            cell.alignment = _wrap_align(
                horizontal="center" if c_idx in (1, 2, 9, 10, 12, 13, 14, 15, 16, 17, 18, 22, 23, 27) else "left"
            )
            cell.border = _border()

            if row_color and nivel in ("alto",):
                cell.fill = _solid("FFF0F0")
            elif r_idx % 2 == 0 and not row_color:
                cell.fill = _solid(COL_ALT_ROW)

        # Highlight nivel_atencao cell itself
        nivel_cell = ws.cell(row=r_idx, column=27)
        if nivel in NIVEL_COLORS:
            nivel_cell.fill = _solid(NIVEL_COLORS[nivel])
            nivel_cell.font = Font(name="Calibri", bold=True, size=9, color="000000")

        ws.row_dimensions[r_idx].height = 36

    _autofilter(ws)


# ── Sheet: Por Seção ──────────────────────────────────────────────────────────
def _build_por_secao(ws, por_secao: list[dict]):
    headers = ["Seção", "Descrição da Seção", "Páginas/Slides", "Qtd. Itens", "IDs dos Itens", "Observações"]
    _apply_header_row(ws, headers, row=1, height=22)
    _freeze(ws, "A2")

    for r_idx, secao in enumerate(por_secao, start=2):
        ids = ", ".join(secao.get("ids_itens", []))
        pages = ", ".join(str(p) for p in secao.get("paginas_slides", []))

        values = [
            secao.get("secao", ""),
            secao.get("descricao_secao", ""),
            pages,
            len(secao.get("ids_itens", [])),
            ids,
            secao.get("observacoes", ""),
        ]
        for c_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=str(value) if value else "")
            cell.font = _body_font()
            cell.alignment = _wrap_align()
            cell.border = _border()
            if r_idx % 2 == 0:
                cell.fill = _solid(COL_ALT_ROW)

        # Alternate header color per section
        ws.cell(row=r_idx, column=1).font = Font(name="Calibri", bold=True, size=9, color="1F3864")
        ws.row_dimensions[r_idx].height = 28

    _set_col_widths(ws, {1: 28, 2: 45, 3: 22, 4: 10, 5: 50, 6: 40})
    _autofilter(ws)


# ── Sheet: Medições ───────────────────────────────────────────────────────────
def _build_medicoes(ws, medicoes: list[dict]):
    headers = [
        "Item", "ID Item", "Página/Slide",
        "Medida Original", "Unidade Original",
        "Medida em Metros", "Tipo de Medida",
        "Confiabilidade", "Observação",
    ]
    fields = [
        "item", "id_item", "pagina_slide",
        "medida_original", "unidade_original",
        "medida_metros", "tipo_medida",
        "confiabilidade", "observacao_confiabilidade",
    ]
    _apply_header_row(ws, headers, row=1, height=22)
    _freeze(ws, "A2")

    CONF_COLORS = {"alta": COL_BAIXO, "média": COL_MEDIO, "media": COL_MEDIO, "baixa": COL_ALTO}

    for r_idx, med in enumerate(medicoes, start=2):
        for c_idx, field in enumerate(fields, start=1):
            value = med.get(field, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=str(value) if value else "")
            cell.font = _body_font()
            cell.alignment = _wrap_align(horizontal="center" if c_idx in (2, 3, 5, 6, 7, 8) else "left")
            cell.border = _border()
            if r_idx % 2 == 0:
                cell.fill = _solid(COL_ALT_ROW)

        conf_val = str(med.get("confiabilidade", "")).strip().lower()
        conf_cell = ws.cell(row=r_idx, column=8)
        if conf_val in CONF_COLORS:
            conf_cell.fill = _solid(CONF_COLORS[conf_val])
            conf_cell.font = Font(name="Calibri", bold=True, size=9)

        ws.row_dimensions[r_idx].height = 22

    _set_col_widths(ws, {1: 35, 2: 10, 3: 14, 4: 20, 5: 14, 6: 16, 7: 16, 8: 14, 9: 40})
    _autofilter(ws)


# ── Sheet: Pendências ─────────────────────────────────────────────────────────
def _build_pendencias(ws, pendencias: list[dict]):
    headers = [
        "ID Item", "Item", "Página/Slide", "Seção",
        "Descrição da Pendência", "Tipo de Pendência",
        "Nível de Urgência", "Responsável", "Prazo",
    ]
    fields = [
        "id_item", "item", "pagina_slide", "secao",
        "descricao_pendencia", "tipo_pendencia",
        "nivel_urgencia", "responsavel", "prazo",
    ]
    _apply_header_row(ws, headers, row=1, height=22)
    _freeze(ws, "A2")

    URG_COLORS = {
        "crítico": "C0392B", "critico": "C0392B",
        "alto": COL_ALTO,
        "médio": COL_MEDIO, "medio": COL_MEDIO,
        "baixo": COL_BAIXO,
    }

    for r_idx, pend in enumerate(pendencias, start=2):
        urgencia = str(pend.get("nivel_urgencia", "")).strip().lower()
        bg = URG_COLORS.get(urgencia, None)

        for c_idx, field in enumerate(fields, start=1):
            value = pend.get(field, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=str(value) if value else "")
            cell.font = _body_font()
            cell.alignment = _wrap_align(horizontal="center" if c_idx in (1, 3, 7) else "left")
            cell.border = _border()
            if bg and urgencia in ("crítico", "critico"):
                cell.fill = _solid("FFF5F5")
            elif r_idx % 2 == 0 and not bg:
                cell.fill = _solid(COL_ALT_ROW)

        urg_cell = ws.cell(row=r_idx, column=7)
        if urgencia in URG_COLORS:
            urg_cell.fill = _solid(URG_COLORS[urgencia])
            urg_cell.font = Font(name="Calibri", bold=True, size=9,
                                 color="FFFFFF" if urgencia in ("crítico", "critico") else "000000")

        ws.row_dimensions[r_idx].height = 30

    _set_col_widths(ws, {1: 10, 2: 35, 3: 12, 4: 22, 5: 50, 6: 22, 7: 14, 8: 20, 9: 15})
    _autofilter(ws)


# ── Sheet: Legenda/Status ─────────────────────────────────────────────────────
def _build_legenda(ws, legenda: list[dict]):
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:E1")
    title = ws["A1"]
    title.value = "LEGENDA DE CORES E STATUS"
    title.font = Font(name="Calibri", bold=True, size=13, color=COL_HEADER_FG)
    title.fill = _solid(COL_HEADER_BG)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["Cor / Hex", "Amostra", "Significado", "Fornecedor", "Status Relacionado", "Observações"]
    _apply_header_row(ws, headers, row=2, height=20)

    for r_idx, leg in enumerate(legenda, start=3):
        hex_color = leg.get("cor_hex", "").replace("#", "").strip()

        ws.cell(row=r_idx, column=1, value=f"#{hex_color}" if hex_color else "Não informado")
        ws.cell(row=r_idx, column=1).font = _body_font()
        ws.cell(row=r_idx, column=1).border = _border()

        # Color sample cell
        sample_cell = ws.cell(row=r_idx, column=2, value="")
        if hex_color and len(hex_color) == 6:
            try:
                sample_cell.fill = _solid(hex_color)
            except Exception:
                pass
        sample_cell.border = _border()

        for c_idx, key in enumerate(["significado", "fornecedor", "status_relacionado", "observacoes"], start=3):
            cell = ws.cell(row=r_idx, column=c_idx, value=str(leg.get(key, "")))
            cell.font = _body_font()
            cell.alignment = _wrap_align()
            cell.border = _border()
            if r_idx % 2 == 0:
                cell.fill = _solid(COL_ALT_ROW)

        ws.row_dimensions[r_idx].height = 22

    # Built-in legend for nivel_atencao
    row = len(legenda) + 5
    ws.merge_cells(f"A{row}:F{row}")
    label = ws.cell(row=row, column=1, value="LEGENDA INTERNA — NÍVEL DE ATENÇÃO / URGÊNCIA")
    label.font = Font(name="Calibri", bold=True, size=10, color=COL_HEADER_FG)
    label.fill = _solid(COL_SECTION_BG)
    label.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1

    built_in = [
        (COL_BAIXO,  "Baixo / Alta confiabilidade",   "Verde",  "Item sem pendências relevantes"),
        (COL_MEDIO,  "Médio / Média confiabilidade",   "Amarelo", "Item com pendências controladas"),
        (COL_ALTO,   "Alto / Baixa confiabilidade",    "Vermelho", "Item crítico ou pendência grave"),
        ("C0392B",   "Crítico",                        "Vermelho escuro", "Pendência bloqueante ou urgente"),
    ]
    for hex_c, nivel, cor_nome, obs in built_in:
        ws.cell(row=row, column=1, value=f"#{hex_c}").font = _body_font()
        ws.cell(row=row, column=1).border = _border()
        ws.cell(row=row, column=2).fill = _solid(hex_c)
        ws.cell(row=row, column=2).border = _border()
        ws.cell(row=row, column=3, value=nivel).font = _body_font(bold=True)
        ws.cell(row=row, column=3).border = _border()
        ws.cell(row=row, column=4, value=cor_nome).font = _body_font()
        ws.cell(row=row, column=4).border = _border()
        ws.cell(row=row, column=5, value=obs).font = _body_font()
        ws.cell(row=row, column=5).border = _border()
        ws.cell(row=row, column=5).alignment = _wrap_align()
        ws.row_dimensions[row].height = 20
        row += 1

    _set_col_widths(ws, {1: 14, 2: 10, 3: 30, 4: 25, 5: 28, 6: 40})


# ── Main builder ─────────────────────────────────────────────────────────────
def generate_excel(data: dict, filename: str) -> bytes:
    """
    Build the full Excel workbook from analyzed data and return as bytes.

    data: output dict from claude_analyzer.analyze_document()
    filename: original uploaded filename (for display in Resumo)
    """
    wb = Workbook()

    # Remove default sheet
    default_ws = wb.active
    wb.remove(default_ws)

    # 1. Resumo
    ws_resumo = wb.create_sheet("Resumo")
    _build_resumo(ws_resumo, data.get("resumo", {}), filename)

    # 2. Itens Detalhados
    ws_itens = wb.create_sheet("Itens Detalhados")
    _build_itens(ws_itens, data.get("itens_detalhados", []))

    # 3. Por Seção
    ws_secao = wb.create_sheet("Por Seção")
    _build_por_secao(ws_secao, data.get("por_secao", []))

    # 4. Medições
    ws_med = wb.create_sheet("Medições")
    _build_medicoes(ws_med, data.get("medicoes", []))

    # 5. Pendências
    ws_pend = wb.create_sheet("Pendências")
    _build_pendencias(ws_pend, data.get("pendencias", []))

    # 6. Legenda / Status
    ws_leg = wb.create_sheet("Legenda e Status")
    _build_legenda(ws_leg, data.get("legenda_status", []))

    # Tab colors
    ws_resumo.sheet_properties.tabColor    = "1F3864"
    ws_itens.sheet_properties.tabColor     = "2E75B6"
    ws_secao.sheet_properties.tabColor     = "4472C4"
    ws_med.sheet_properties.tabColor       = "70AD47"
    ws_pend.sheet_properties.tabColor      = "FF6B6B"
    ws_leg.sheet_properties.tabColor       = "FFC000"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
