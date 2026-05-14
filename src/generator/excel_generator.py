"""
Excel generator for decupagem output — simplified supplier-first layout.
Sheets:
  1. Resumo
  2. Todos os Itens
  3+. One tab per supplier (Guedes Cenografia, MGTT, Ricardo, Outros, Sem Fornecedor)
  Last. Pendências
"""
from __future__ import annotations

import re
import io
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Palette ───────────────────────────────────────────────────────────────────
_NAV   = "1F3864"   # dark navy — main headers
_BLU   = "2E75B6"   # mid-blue — sub-headers
_ALT   = "EBF3FB"   # light-blue alternating row
_ALTO  = "FF6B6B"   # high priority
_MEDIO = "FFD93D"   # medium
_BAIXO = "C6EFCE"   # low / ok
_WHITE = "FFFFFF"
_GRAY  = "F2F2F2"
_BORD  = "BDC3C7"

# Supplier tab colors (cycling)
_SUP_COLORS = ["2E75B6", "70AD47", "ED7D31", "9E480E", "7030A0", "00B0F0", "FF0000"]

NIVEL_FILL = {
    "alto":   _ALTO,
    "médio":  _MEDIO,
    "medio":  _MEDIO,
    "baixo":  _BAIXO,
}

# Priority order for supplier tabs
_PRIORITY_SUPPLIERS = ["Guedes Cenografia", "MGTT", "Ricardo"]


# ── Style helpers ─────────────────────────────────────────────────────────────

def _solid(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _border(color=_BORD, style="thin") -> Border:
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _font(bold=False, size=9, color="000000") -> Font:
    return Font(name="Calibri", bold=bold, size=size, color=color)


def _align(h="left", v="center", wrap=True) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _hdr(ws, row: int, col: int, value: str, bg=_NAV, fg=_WHITE, size=10, bold=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="Calibri", bold=bold, size=size, color=fg)
    c.fill      = _solid(bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = _border()
    return c


def _cell(ws, row: int, col: int, value, bg=None, bold=False, h="left", size=9):
    c = ws.cell(row=row, column=col, value=str(value) if value is not None else "")
    c.font      = _font(bold=bold, size=size)
    c.alignment = _align(h=h)
    c.border    = _border()
    if bg:
        c.fill = _solid(bg)
    return c


def _widths(ws, mapping: dict[int, float]):
    for col, w in mapping.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def _sanitize_tab(name: str) -> str:
    cleaned = re.sub(r'[\\/:?*\[\]]+', ' ', name).strip()
    return cleaned[:31]


# ── Measurement formatter ─────────────────────────────────────────────────────

def _fmt_med(item: dict) -> str:
    """Return a human-readable measurement string like '3,00m × 2,50m × 0,50m'."""
    orig = (item.get("medida_original") or "").strip()
    if orig:
        return orig

    parts = []
    for key, label in [
        ("largura", "L"), ("altura", "A"), ("profundidade", "P"),
        ("comprimento", "C"), ("espessura", "esp"), ("area", "área"),
        ("metro_linear", "ml"),
    ]:
        val = (item.get(key) or "").strip()
        if val and val not in ("A confirmar", ""):
            # Normalize: try to add 'm' if it looks like a number
            try:
                fv = float(val.replace(",", "."))
                parts.append(f"{fv:.2f}m".replace(".", ","))
            except ValueError:
                parts.append(f"{label}: {val}")
    return " × ".join(parts) if parts else "—"


def _fmt_med_detail(item: dict) -> list[tuple[str, str]]:
    """Return list of (dimensão, valor) for each filled measurement."""
    rows = []
    for key, label in [
        ("largura", "Largura"), ("altura", "Altura"), ("profundidade", "Profundidade"),
        ("comprimento", "Comprimento"), ("espessura", "Espessura"),
        ("area", "Área"), ("metro_linear", "Metro Linear"),
    ]:
        val = (item.get(key) or "").strip()
        if val and val not in ("A confirmar", ""):
            try:
                fv = float(val.replace(",", "."))
                rows.append((label, f"{fv:.2f}m".replace(".", ",")))
            except ValueError:
                rows.append((label, val))
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 1. RESUMO
# ─────────────────────────────────────────────────────────────────────────────

def _build_resumo(ws, data: dict, filename: str):
    ws.sheet_view.showGridLines = False
    resumo   = data.get("resumo", {})
    itens    = data.get("itens_detalhados", [])
    n_pend   = len(data.get("pendencias", []))

    # Count per supplier
    sup_counts: dict[str, int] = defaultdict(int)
    for it in itens:
        forn = (it.get("fornecedor") or "").strip()
        sup_counts[forn or "Sem Fornecedor"] += 1

    # Title
    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value     = "DECUPAGEM TÉCNICA — RESUMO"
    t.font      = Font(name="Calibri", bold=True, size=14, color=_WHITE)
    t.fill      = _solid(_NAV)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Info rows
    fields = [
        ("Arquivo",                filename),
        ("Total de páginas/slides", resumo.get("total_paginas_slides", "—")),
        ("Total de itens",          len(itens)),
        ("Total de pendências",     n_pend),
    ]
    for i, (k, v) in enumerate(fields):
        r = i + 2
        c1 = ws.cell(row=r, column=1, value=k)
        c1.font = _font(bold=True, size=9, color=_WHITE)
        c1.fill = _solid(_BLU)
        c1.alignment = _align(h="left")
        c1.border = _border()

        c2 = ws.cell(row=r, column=2, value=str(v))
        c2.font = _font(size=9)
        c2.alignment = _align(h="left")
        c2.border = _border()

    # Supplier breakdown header
    r_start = len(fields) + 3
    ws.merge_cells(f"A{r_start}:D{r_start}")
    hrow = ws[f"A{r_start}"]
    hrow.value     = "ITENS POR FORNECEDOR"
    hrow.font      = Font(name="Calibri", bold=True, size=11, color=_WHITE)
    hrow.fill      = _solid(_NAV)
    hrow.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r_start].height = 22

    _hdr(ws, r_start + 1, 1, "Fornecedor", bg=_BLU)
    _hdr(ws, r_start + 1, 2, "Qtd. Itens", bg=_BLU)

    for i, (sup, cnt) in enumerate(sorted(sup_counts.items(), key=lambda x: -x[1])):
        r = r_start + 2 + i
        bg = _ALT if i % 2 == 0 else None
        _cell(ws, r, 1, sup, bg=bg, bold=True)
        _cell(ws, r, 2, cnt, bg=bg, h="center")

    _widths(ws, {1: 30, 2: 15, 3: 20, 4: 20})


# ─────────────────────────────────────────────────────────────────────────────
# 2. TODOS OS ITENS
# ─────────────────────────────────────────────────────────────────────────────

# Columns for item sheets
_ITEM_COLS = [
    ("ID",           "id",                          6),
    ("Fornecedor",   "fornecedor",                  20),
    ("Seção",        "secao",                       18),
    ("Item",         "item",                        40),
    ("Qtd",          "quantidade",                   7),
    ("Un",           "unidade",                      6),
    ("Medidas",      "__medidas__",                 25),
    ("Categoria",    "categoria",                   18),
    ("Material",     "material",                    18),
    ("Cor/Acabamento","acabamento_cor",             15),
    ("Tipo Produção","tipo_producao",               15),
    ("Status Arte",  "status_arte",                 14),
    ("Status Compra","status_compra_locacao",       14),
    ("Pendência",    "pendencia_acao_necessaria",   30),
    ("Nível",        "nivel_atencao",               10),
    ("Obs. Técnicas","observacoes_tecnicas",        30),
    ("Página/Slide", "pagina_slide_origem",          8),
]


def _write_item_headers(ws, row=1, bg=_NAV):
    for c_idx, (label, _, _w) in enumerate(_ITEM_COLS, 1):
        _hdr(ws, row, c_idx, label, bg=bg)
    ws.freeze_panes = ws.cell(row=2, column=1)
    ws.auto_filter.ref = ws.dimensions
    _widths(ws, {c: w for c, (_, _, w) in enumerate(_ITEM_COLS, 1)})


def _write_item_row(ws, row: int, item: dict, alt: bool):
    nivel = (item.get("nivel_atencao") or "").lower()
    row_bg = NIVEL_FILL.get(nivel) if nivel in ("alto",) else (_ALT if alt else None)

    for c_idx, (_, key, _) in enumerate(_ITEM_COLS, 1):
        if key == "__medidas__":
            val = _fmt_med(item)
        else:
            val = item.get(key, "") or ""

        bg = row_bg
        if key == "nivel_atencao" and nivel:
            bg = NIVEL_FILL.get(nivel, row_bg)

        _cell(ws, row, c_idx, val, bg=bg)

    ws.row_dimensions[row].height = 18


def _build_todos_itens(ws, itens: list[dict]):
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:Q1")
    t = ws["A1"]
    t.value     = "TODOS OS ITENS — DECUPAGEM COMPLETA"
    t.font      = Font(name="Calibri", bold=True, size=12, color=_WHITE)
    t.fill      = _solid(_NAV)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    _write_item_headers(ws, row=2)

    for i, item in enumerate(itens):
        _write_item_row(ws, i + 3, item, alt=i % 2 == 0)


# ─────────────────────────────────────────────────────────────────────────────
# 3. SUPPLIER SHEET
# ─────────────────────────────────────────────────────────────────────────────

# Supplier-sheet columns — fewer, more focused
_SUP_COLS = [
    ("ID",           "id",                           6),
    ("Seção",        "secao",                        18),
    ("Item",         "item",                         42),
    ("Qtd",          "quantidade",                    7),
    ("Un",           "unidade",                       6),
    ("Medidas",      "__medidas__",                  28),
    ("Largura",      "__L__",                        10),
    ("Altura",       "__A__",                        10),
    ("Profundidade", "__P__",                        12),
    ("Comprimento",  "__C__",                        12),
    ("Categoria",    "categoria",                    18),
    ("Material",     "material",                     18),
    ("Cor/Acabamento","acabamento_cor",              14),
    ("Tipo Produção","tipo_producao",                15),
    ("Status Compra","status_compra_locacao",        14),
    ("Pendência",    "pendencia_acao_necessaria",    30),
    ("Nível",        "nivel_atencao",                10),
    ("Obs",          "observacoes_tecnicas",         25),
    ("Pág/Slide",    "pagina_slide_origem",           8),
]


def _fmt_dim(item: dict, dim: str) -> str:
    val = (item.get(dim) or "").strip()
    if not val or val == "A confirmar":
        return "—"
    try:
        return f"{float(val.replace(',', '.')):.2f}m".replace(".", ",")
    except ValueError:
        return val


def _build_supplier_sheet(ws, supplier_name: str, itens: list[dict], tab_color: str):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color

    # Title row
    n_cols = len(_SUP_COLS)
    title_cell = f"A1:{get_column_letter(n_cols)}1"
    ws.merge_cells(title_cell)
    t = ws["A1"]
    t.value     = f"FORNECEDOR: {supplier_name.upper()}  ({len(itens)} itens)"
    t.font      = Font(name="Calibri", bold=True, size=12, color=_WHITE)
    t.fill      = _solid(_NAV)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Headers
    for c_idx, (label, _, _) in enumerate(_SUP_COLS, 1):
        _hdr(ws, 2, c_idx, label, bg=_BLU)
    ws.freeze_panes = ws.cell(row=3, column=1)
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"
    _widths(ws, {c: w for c, (_, _, w) in enumerate(_SUP_COLS, 1)})

    # Data rows
    for i, item in enumerate(itens):
        r   = i + 3
        nivel = (item.get("nivel_atencao") or "").lower()
        alt   = i % 2 == 0
        row_bg = NIVEL_FILL.get(nivel) if nivel == "alto" else (_ALT if alt else None)

        for c_idx, (_, key, _) in enumerate(_SUP_COLS, 1):
            if key == "__medidas__":
                val = _fmt_med(item)
            elif key == "__L__":
                val = _fmt_dim(item, "largura")
            elif key == "__A__":
                val = _fmt_dim(item, "altura")
            elif key == "__P__":
                val = _fmt_dim(item, "profundidade")
            elif key == "__C__":
                val = _fmt_dim(item, "comprimento")
            else:
                val = item.get(key, "") or ""

            bg = row_bg
            if key == "nivel_atencao" and nivel:
                bg = NIVEL_FILL.get(nivel, row_bg)
            _cell(ws, r, c_idx, val, bg=bg)

        ws.row_dimensions[r].height = 18


# ─────────────────────────────────────────────────────────────────────────────
# 4. PENDÊNCIAS
# ─────────────────────────────────────────────────────────────────────────────

_PEND_COLS = [
    ("ID",              "id",                          6),
    ("Fornecedor",      "fornecedor",                 20),
    ("Seção",           "secao",                      16),
    ("Item",            "item",                       40),
    ("Pendência",       "pendencia_acao_necessaria",  35),
    ("Nível",           "nivel_atencao",              10),
    ("Status Arte",     "status_arte",                14),
    ("Status Compra",   "status_compra_locacao",      14),
    ("Pág/Slide",       "pagina_slide_origem",          8),
]


def _build_pendencias(ws, itens: list[dict]):
    ws.sheet_view.showGridLines = False
    pend_items = [it for it in itens if it.get("pendencia_acao_necessaria", "").strip()]

    ws.merge_cells(f"A1:{get_column_letter(len(_PEND_COLS))}1")
    t = ws["A1"]
    t.value     = f"PENDÊNCIAS E AÇÕES NECESSÁRIAS  ({len(pend_items)} itens)"
    t.font      = Font(name="Calibri", bold=True, size=12, color=_WHITE)
    t.fill      = _solid(_ALTO)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for c_idx, (label, _, _) in enumerate(_PEND_COLS, 1):
        _hdr(ws, 2, c_idx, label, bg="C0392B", fg=_WHITE)
    ws.freeze_panes = ws.cell(row=3, column=1)
    ws.auto_filter.ref = f"A2:{get_column_letter(len(_PEND_COLS))}2"
    _widths(ws, {c: w for c, (_, _, w) in enumerate(_PEND_COLS, 1)})

    for i, item in enumerate(pend_items):
        r   = i + 3
        nivel = (item.get("nivel_atencao") or "").lower()
        bg = NIVEL_FILL.get(nivel) if nivel in NIVEL_FILL else (_ALT if i % 2 == 0 else None)
        for c_idx, (_, key, _) in enumerate(_PEND_COLS, 1):
            _cell(ws, r, c_idx, item.get(key, "") or "", bg=bg)
        ws.row_dimensions[r].height = 18


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def generate_excel(data: dict, filename: str = "documento") -> bytes:
    """
    data: output dict from analyze_document()
    Returns raw .xlsx bytes.
    """
    wb = Workbook()
    itens: list[dict] = data.get("itens_detalhados", [])

    # ── 1. Resumo ──────────────────────────────────────────────────────────────
    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    ws_resumo.sheet_properties.tabColor = _NAV
    _build_resumo(ws_resumo, data, filename)

    # ── 2. Todos os Itens ──────────────────────────────────────────────────────
    ws_all = wb.create_sheet("Todos os Itens")
    ws_all.sheet_properties.tabColor = _BLU
    _build_todos_itens(ws_all, itens)

    # ── 3. Supplier sheets ─────────────────────────────────────────────────────
    # Bucket items by supplier
    buckets: dict[str, list[dict]] = defaultdict(list)
    sem_forn: list[dict] = []

    for item in itens:
        forn = (item.get("fornecedor") or "").strip()
        if not forn or forn.lower() in ("não informado", "nao informado", "a confirmar"):
            sem_forn.append(item)
        else:
            buckets[forn].append(item)

    # Determine tab order: priority suppliers first, then alphabetical others
    outros_items: list[dict] = []
    ordered_suppliers: list[str] = []
    for ps in _PRIORITY_SUPPLIERS:
        if ps in buckets:
            ordered_suppliers.append(ps)
    for name in sorted(buckets.keys()):
        if name not in _PRIORITY_SUPPLIERS:
            outros_items.extend(buckets[name])

    for s_idx, sup_name in enumerate(ordered_suppliers):
        tab_color = _SUP_COLORS[s_idx % len(_SUP_COLORS)]
        ws_s = wb.create_sheet(_sanitize_tab(sup_name))
        ws_s.sheet_properties.tabColor = tab_color
        _build_supplier_sheet(ws_s, sup_name, buckets[sup_name], tab_color)

    # "Outros Fornecedores" groups remaining known suppliers
    remaining = {k: v for k, v in buckets.items() if k not in _PRIORITY_SUPPLIERS}
    if remaining:
        outros_all: list[dict] = []
        for items_list in remaining.values():
            outros_all.extend(items_list)
        ws_out = wb.create_sheet("Outros Fornecedores")
        ws_out.sheet_properties.tabColor = _SUP_COLORS[3 % len(_SUP_COLORS)]
        _build_supplier_sheet(ws_out, "Outros Fornecedores", outros_all, _SUP_COLORS[3 % len(_SUP_COLORS)])

    # "Sem Fornecedor"
    if sem_forn:
        ws_no = wb.create_sheet("Sem Fornecedor")
        ws_no.sheet_properties.tabColor = "808080"
        _build_supplier_sheet(ws_no, "Sem Fornecedor / A Definir", sem_forn, "808080")

    # ── 4. Pendências ──────────────────────────────────────────────────────────
    ws_pend = wb.create_sheet("Pendências")
    ws_pend.sheet_properties.tabColor = "FF0000"
    _build_pendencias(ws_pend, itens)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
