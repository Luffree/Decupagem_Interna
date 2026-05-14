"""
Analisador local de documentos — sem API externa.
Processa o conteúdo extraído de PDFs/PPTXs usando regex, correspondência
de palavras-chave e heurísticas de projeto técnico em português.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from .patterns import (
    MEASURE_PATTERNS_LIST, CATEGORY_MAP, MATERIAL_KEYWORDS,
    COLOR_KEYWORDS, ACABAMENTO_KEYWORDS, TIPO_PRODUCAO_MAP,
    STATUS_ARTE_KEYWORDS, STATUS_COMPRA_KEYWORDS, PENDENCIA_KEYWORDS,
    NIVEL_ALTO_KEYWORDS, NIVEL_MEDIO_KEYWORDS, TABLE_HEADER_SYNONYMS,
    SECTION_INDICATOR_WORDS,
)


# ─────────────────────────────────────────────────────────────────────────────
# Public entry point (same signature as the old claude_analyzer)
# ─────────────────────────────────────────────────────────────────────────────

def analyze_document(doc_dict: dict, doc_type: str, progress_callback=None, **kwargs) -> dict:
    """
    Analisa o conteúdo extraído e retorna o dicionário estruturado
    esperado pelo excel_generator.
    doc_type: "pdf" ou "pptx"
    """
    analyzer = _DocumentAnalyzer(doc_dict, doc_type, progress_callback)
    return analyzer.run()


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    """Lower-case + collapse whitespace."""
    return re.sub(r'\s+', ' ', str(text).lower()).strip()


def _clean_cell(v) -> str:
    s = str(v).strip() if v is not None else ''
    return '' if s in ('None', 'nan', 'NaN') else s


def _to_meters(value_str: str, unit: str) -> str:
    """Convert a numeric string + unit to meters."""
    try:
        v = float(value_str.replace(',', '.'))
        u = unit.lower().strip()
        if u == 'mm':
            return f"{v / 1000:.4f}"
        if u == 'cm':
            return f"{v / 100:.3f}"
        if u in ('m', 'metro', 'metros'):
            return f"{v:.3f}"
    except (ValueError, AttributeError):
        pass
    return 'A confirmar'


# ─────────────────────────────────────────────────────────────────────────────
# Main analyzer class
# ─────────────────────────────────────────────────────────────────────────────

class _DocumentAnalyzer:
    def __init__(self, doc_dict: dict, doc_type: str, progress_callback=None):
        self.doc = doc_dict
        self.doc_type = doc_type
        self.cb = progress_callback or (lambda msg: None)
        self._counter = 0
        self.items: list[dict] = []

        if doc_type == 'pdf':
            self._units_key = 'pages'
            self._num_key   = 'page_number'
            self._total_key = 'total_pages'
            self._unit_label = 'Página'
        else:
            self._units_key = 'slides'
            self._num_key   = 'slide_number'
            self._total_key = 'total_slides'
            self._unit_label = 'Slide'

    # ── Run ──────────────────────────────────────────────────────────────────

    def run(self) -> dict:
        self.cb("Extraindo metadados do arquivo...")
        self._meta = self._read_metadata()

        units = self.doc.get(self._units_key, [])
        total = len(units)

        for idx, unit in enumerate(units):
            pct = int(30 + (idx / max(total, 1)) * 55)
            page_num = str(unit.get(self._num_key, idx + 1))
            title = _clean_cell(unit.get('title', ''))
            section = title or f"{self._unit_label} {page_num}"
            self.cb(f"Analisando {self._unit_label} {page_num}/{total}: {section[:40]}...")

            # Tables (highest fidelity)
            for table in unit.get('tables', []):
                self._parse_table(table, page_num, section)

            # Main text
            main_text = _clean_cell(unit.get('text', ''))
            if main_text:
                self._parse_text(main_text, page_num, section, 'Texto principal')

            # Text blocks (PPTX)
            for block in unit.get('text_blocks', []):
                self._parse_text(_clean_cell(block), page_num, section, 'Bloco de texto')

            # Notes (PPTX)
            notes = _clean_cell(unit.get('notes', ''))
            if notes:
                self._parse_text(notes, page_num, section, 'Notas do slide')

            # Shapes: only process when they carry a fill_color with semantic meaning.
            # Text-only shapes are already covered by text_blocks; processing them
            # again here would create duplicates.
            for sh in unit.get('shapes_summary', []):
                fill_color = (sh.get('fill_color') or '').strip()
                sh_text    = _clean_cell(sh.get('text', ''))
                if fill_color and len(fill_color) == 6 and len(sh_text) > 3:
                    self._parse_shape(sh_text, page_num, section, fill_color)

            # PDF annotations
            for ann in unit.get('annotations', []):
                ann_text = _clean_cell(ann.get('content', ''))
                if len(ann_text) > 3:
                    self._parse_text(ann_text, page_num, section,
                                     f"Anotação ({ann.get('type', '?')})")

        self.cb("Consolidando itens e removendo duplicatas...")
        self._deduplicate()

        self.cb("Montando estrutura final...")
        return self._build_output()

    # ── Metadata ─────────────────────────────────────────────────────────────

    def _read_metadata(self) -> dict:
        meta = self.doc.get('metadata', {}) or {}
        filename = self.doc.get('filename', '') or ''
        stem = Path(filename).stem if filename else ''

        client = (meta.get('author') or meta.get('creator') or
                  meta.get('Author') or 'Não informado').strip()
        project = (meta.get('title') or meta.get('subject') or
                   meta.get('Title') or stem or 'Não informado').strip()
        return {'client': client, 'project': project}

    # ── ID factory ───────────────────────────────────────────────────────────

    def _next_id(self, section: str = '') -> str:
        self._counter += 1
        prefix = re.sub(r'[^A-Z0-9]', '', section.upper())[:4] or 'IT'
        return f"{prefix}-{self._counter:04d}"

    # ── Table parsing ─────────────────────────────────────────────────────────

    def _parse_table(self, table: list[list], page_num: str, section: str):
        if not table:
            return

        # Detect header row
        header_row: Optional[list] = None
        col_map: dict[str, int] = {}
        data_start = 0

        if len(table) >= 2:
            first = [_norm(c) for c in table[0]]
            score = sum(
                1 for cell in first
                if any(syn in cell for syns in TABLE_HEADER_SYNONYMS.values() for syn in syns)
            )
            if score >= 2 or (score >= 1 and len(first) >= 3):
                header_row = table[0]
                col_map = self._map_columns(header_row)
                data_start = 1

        for row in table[data_start:]:
            cells = [_clean_cell(c) for c in row]
            if not any(cells):
                continue

            item = self._row_to_item(cells, col_map, page_num, section)
            if item:
                self.items.append(item)

    def _map_columns(self, header_row: list) -> dict[str, int]:
        col_map: dict[str, int] = {}
        for idx, cell in enumerate(header_row):
            cell_norm = _norm(cell)
            for field, synonyms in TABLE_HEADER_SYNONYMS.items():
                if field not in col_map and any(syn in cell_norm for syn in synonyms):
                    col_map[field] = idx
        return col_map

    def _row_to_item(self, cells: list[str], col_map: dict, page_num: str, section: str) -> Optional[dict]:
        def gcol(field: str) -> str:
            idx = col_map.get(field)
            return cells[idx] if idx is not None and idx < len(cells) else ''

        # Determine item name
        item_name = gcol('item') or gcol('descricao') or gcol('nome') or gcol('codigo')
        if not item_name:
            item_name = next((c for c in cells if c), '')
        if not item_name:
            return None

        all_text = ' | '.join(c for c in cells if c)
        qty       = gcol('quantidade')
        unidade   = gcol('unidade')
        med_raw   = gcol('medida')
        material  = gcol('material')
        status    = gcol('status')
        obs       = gcol('observacao')
        fornecedor = gcol('fornecedor')
        categoria  = gcol('categoria')
        ambiente   = gcol('ambiente')

        detected  = self._classify(all_text + ' ' + item_name)
        meds      = self._extract_meds(all_text + ' ' + med_raw)
        st_arte   = self._detect_status_arte(all_text + ' ' + status)
        st_compra = self._detect_status_compra(all_text + ' ' + status)
        pend      = self._detect_pendencia(all_text)
        nivel     = self._detect_nivel(all_text, pend)

        return {
            'id':                       self._next_id(section),
            'pagina_slide_origem':      page_num,
            'secao':                    section,
            'sub_secao_ambiente':       ambiente or '',
            'fornecedor':               fornecedor or detected.get('fornecedor', 'Não informado'),
            'categoria':                categoria or detected.get('categoria', 'A confirmar'),
            'item':                     item_name[:120],
            'descricao_detalhada':      all_text if all_text != item_name else '',
            'quantidade':               qty or detected.get('quantidade', 'A confirmar'),
            'unidade':                  unidade or detected.get('unidade', ''),
            'medida_original':          meds.get('medida_original', '') or med_raw,
            'largura':                  meds.get('largura', ''),
            'altura':                   meds.get('altura', ''),
            'profundidade':             meds.get('profundidade', ''),
            'comprimento':              meds.get('comprimento', ''),
            'espessura':                meds.get('espessura', ''),
            'area':                     meds.get('area', ''),
            'metro_linear':             meds.get('metro_linear', ''),
            'material':                 material or detected.get('material', ''),
            'acabamento_cor':           detected.get('acabamento', ''),
            'tipo_producao':            detected.get('tipo_producao', 'A confirmar'),
            'status_arte':              st_arte or 'A confirmar',
            'status_compra_locacao':    status or st_compra or 'A confirmar',
            'responsavel_area':         'Não informado',
            'observacoes_tecnicas':     obs,
            'pendencia_acao_necessaria': pend,
            'nivel_atencao':            nivel,
            'fonte_informacao':         'Tabela',
        }

    # ── Text parsing ──────────────────────────────────────────────────────────

    def _parse_text(self, text: str, page_num: str, section: str, source: str):
        if not text.strip():
            return

        lines = [l.strip() for l in re.split(r'[\n\r]+', text) if l.strip()]
        current_sub = ''

        for line in lines:
            if len(line) < 4:
                continue

            if self._is_header(line):
                current_sub = re.sub(r':$', '', line).strip()
                continue

            item_text = self._strip_bullet(line)
            if not item_text or len(item_text) < 4:
                continue

            # Skip lines that are only numbers or punctuation
            if re.match(r'^[\d\s\.\,\-\:\/]+$', item_text):
                continue

            detected  = self._classify(item_text)
            meds      = self._extract_meds(item_text)
            st_arte   = self._detect_status_arte(item_text)
            st_compra = self._detect_status_compra(item_text)
            pend      = self._detect_pendencia(item_text)
            nivel     = self._detect_nivel(item_text, pend)

            self.items.append({
                'id':                       self._next_id(section),
                'pagina_slide_origem':      page_num,
                'secao':                    section,
                'sub_secao_ambiente':       current_sub,
                'fornecedor':               detected.get('fornecedor', 'Não informado'),
                'categoria':                detected.get('categoria', 'A confirmar'),
                'item':                     item_text[:120],
                'descricao_detalhada':      item_text,
                'quantidade':               detected.get('quantidade', 'A confirmar'),
                'unidade':                  detected.get('unidade', ''),
                'medida_original':          meds.get('medida_original', ''),
                'largura':                  meds.get('largura', ''),
                'altura':                   meds.get('altura', ''),
                'profundidade':             meds.get('profundidade', ''),
                'comprimento':              meds.get('comprimento', ''),
                'espessura':                meds.get('espessura', ''),
                'area':                     meds.get('area', ''),
                'metro_linear':             meds.get('metro_linear', ''),
                'material':                 detected.get('material', ''),
                'acabamento_cor':           detected.get('acabamento', ''),
                'tipo_producao':            detected.get('tipo_producao', ''),
                'status_arte':              st_arte,
                'status_compra_locacao':    st_compra,
                'responsavel_area':         'Não informado',
                'observacoes_tecnicas':     '',
                'pendencia_acao_necessaria': pend,
                'nivel_atencao':            nivel,
                'fonte_informacao':         source,
            })

    def _parse_shape(self, text: str, page_num: str, section: str, fill_color: str):
        if len(text) < 4:
            return
        detected  = self._classify(text)
        meds      = self._extract_meds(text)
        pend      = self._detect_pendencia(text)
        nivel     = self._detect_nivel(text, pend)

        color_note = ''
        if fill_color and len(fill_color) == 6:
            color_note = f"Forma com preenchimento #{fill_color}"

        self.items.append({
            'id':                       self._next_id(section),
            'pagina_slide_origem':      page_num,
            'secao':                    section,
            'sub_secao_ambiente':       '',
            'fornecedor':               'Não informado',
            'categoria':                detected.get('categoria', 'A confirmar'),
            'item':                     text[:120],
            'descricao_detalhada':      text,
            'quantidade':               'A confirmar',
            'unidade':                  '',
            'medida_original':          meds.get('medida_original', ''),
            'largura':                  meds.get('largura', ''),
            'altura':                   meds.get('altura', ''),
            'profundidade':             meds.get('profundidade', ''),
            'comprimento':              meds.get('comprimento', ''),
            'espessura':                meds.get('espessura', ''),
            'area':                     meds.get('area', ''),
            'metro_linear':             meds.get('metro_linear', ''),
            'material':                 detected.get('material', ''),
            'acabamento_cor':           (f"#{fill_color} " if fill_color and len(fill_color) == 6 else '') + detected.get('acabamento', ''),
            'tipo_producao':            detected.get('tipo_producao', ''),
            'status_arte':              '',
            'status_compra_locacao':    '',
            'responsavel_area':         'Não informado',
            'observacoes_tecnicas':     color_note,
            'pendencia_acao_necessaria': pend,
            'nivel_atencao':            nivel,
            'fonte_informacao':         'Forma visual',
        })

    # ── Classification helpers ────────────────────────────────────────────────

    def _classify(self, text: str) -> dict:
        tl = _norm(text)
        result: dict = {}

        # Supplier — look for explicit "Fornecedor: X" / "exec.: X" patterns
        forn_m = re.search(
            r'(?:fornecedor|exec\.?|executante|responsável|empresa)\s*[:\-]\s*([^|\n,;—]{3,50})',
            text, re.IGNORECASE,
        )
        if forn_m:
            result['fornecedor'] = forn_m.group(1).strip().rstrip('.')

        # Category
        for category, keywords in CATEGORY_MAP.items():
            if any(kw in tl for kw in keywords):
                result['categoria'] = category
                break
        result.setdefault('categoria', 'A confirmar')

        # Material
        mats = [mat for mat, kws in MATERIAL_KEYWORDS.items() if any(k in tl for k in kws)]
        result['material'] = ', '.join(mats) if mats else ''

        # Color + acabamento
        colors = [c for c in COLOR_KEYWORDS if c.lower() in tl]
        acab   = [a for a in ACABAMENTO_KEYWORDS if a.lower() in tl]
        result['acabamento'] = ', '.join(dict.fromkeys(colors + acab))  # preserve order, dedup

        # Tipo de produção
        for tipo, kws in TIPO_PRODUCAO_MAP.items():
            if any(k in tl for k in kws):
                result['tipo_producao'] = tipo
                break

        # Quantity + unit inline (e.g. "3 un", "2 peças", "10 metros")
        m = re.search(
            r'\b(\d+)\s*(un(?:idades?)?|pçs?|peças?|conjuntos?|jogos?|kits?|metros?|m\b|cm|mm)\b',
            tl,
        )
        if m:
            result['quantidade'] = m.group(1)
            result['unidade']    = m.group(2)

        return result

    def _extract_meds(self, text: str) -> dict:
        """Find first matching measurement pattern and return structured fields."""
        tn = text.replace(',', '.')
        result: dict = {}

        for pattern, mtype, fields in MEASURE_PATTERNS_LIST:
            m = re.search(pattern, tn, re.IGNORECASE)
            if not m:
                continue

            result['medida_original'] = m.group(0).replace('.', ',')
            groups = m.groups()

            if mtype == 'LxAxP' and len(groups) >= 3:
                result['largura']      = groups[0]
                result['altura']       = groups[1]
                result['profundidade'] = groups[2]
                unit = groups[3] if len(groups) > 3 else ''
                result['_unit'] = unit
            elif mtype == 'LxA' and len(groups) >= 2:
                result['largura'] = groups[0]
                result['altura']  = groups[1]
                unit = groups[2] if len(groups) > 2 else ''
                result['_unit'] = unit
            elif mtype == 'area' and groups:
                result['area'] = groups[0]
            elif mtype == 'ml' and groups:
                result['metro_linear'] = groups[0]
            elif mtype in ('espessura',) and groups:
                result['espessura'] = groups[0]
            elif mtype in ('simples', 'diametro') and groups:
                result['comprimento'] = groups[0]
                result['_unit'] = groups[1] if len(groups) > 1 else ''
            break

        return result

    def _detect_status_arte(self, text: str) -> str:
        tl = _norm(text)
        for status, kws in STATUS_ARTE_KEYWORDS.items():
            if any(k in tl for k in kws):
                return status
        return ''

    def _detect_status_compra(self, text: str) -> str:
        tl = _norm(text)
        for status, kws in STATUS_COMPRA_KEYWORDS.items():
            if any(k in tl for k in kws):
                return status
        return ''

    def _detect_pendencia(self, text: str) -> str:
        tl = _norm(text)
        found = []
        for kw in PENDENCIA_KEYWORDS:
            # Short keywords (<=4 chars) must match at word boundaries to avoid
            # false positives (e.g. "nd" matching inside "fundo").
            if len(kw) <= 4:
                if re.search(r'(?<!\w)' + re.escape(kw) + r'(?!\w)', tl):
                    found.append(kw)
            else:
                if kw in tl:
                    found.append(kw)
        if not found:
            return ''
        return 'Verificar: ' + ', '.join(dict.fromkeys(found[:3]))

    def _detect_nivel(self, text: str, pendencia: str) -> str:
        if not pendencia:
            return 'Baixo'
        tl = _norm(text)
        if any(k in tl for k in NIVEL_ALTO_KEYWORDS):
            return 'Alto'
        return 'Médio'

    @staticmethod
    def _is_header(line: str) -> bool:
        s = line.strip()
        if len(s) < 3 or len(s) > 90:
            return False
        if s.isupper() and len(s) > 4:
            return True
        if s.endswith(':') and not re.search(r'[,;]', s):
            return True
        if re.match(r'^\d+[\.\)]\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÜ]', s):
            return True
        return False

    @staticmethod
    def _strip_bullet(line: str) -> str:
        cleaned = re.sub(r'^[\-\•\*\–\—\►\→\·\○\●\□\■\✓\✗\>\♦◆▪▸]+\s*', '', line)
        cleaned = re.sub(r'^\d+[\.\)]\s+', '', cleaned)
        return cleaned.strip()

    # ── Deduplication ─────────────────────────────────────────────────────────

    def _deduplicate(self):
        seen: set = set()
        unique: list = []
        for item in self.items:
            key = (_norm(item['item'])[:60], item['pagina_slide_origem'])
            if key not in seen:
                seen.add(key)
                unique.append(item)
        self.items = unique

    # ── Output builder ────────────────────────────────────────────────────────

    def _build_output(self) -> dict:
        total_units = self.doc.get(self._total_key, 0)

        # ── Seções ────────────────────────────────────────────────────────────
        sec_map: dict[str, dict] = {}
        for item in self.items:
            s = item['secao']
            if s not in sec_map:
                sec_map[s] = {'ids': [], 'pages': set()}
            sec_map[s]['ids'].append(item['id'])
            sec_map[s]['pages'].add(str(item['pagina_slide_origem']))

        por_secao = [
            {
                'secao':           s,
                'descricao_secao': f"Conteúdo identificado em '{s}'",
                'ids_itens':       d['ids'],
                'paginas_slides':  sorted(d['pages']),
                'observacoes':     '',
            }
            for s, d in sec_map.items()
        ]

        # ── Medições ──────────────────────────────────────────────────────────
        medicoes: list[dict] = []
        for item in self.items:
            med = item.get('medida_original', '').strip()
            if not med:
                continue

            tipo = 'medida'
            if item.get('area'):          tipo = 'área'
            elif item.get('metro_linear'): tipo = 'metro linear'
            elif item.get('espessura'):    tipo = 'espessura'
            elif item.get('profundidade'): tipo = 'largura x altura x profundidade'
            elif item.get('largura') and item.get('altura'): tipo = 'largura x altura'

            unit_m = re.search(r'(cm|mm|m\b)', med, re.IGNORECASE)
            unit_str = unit_m.group(1) if unit_m else ''
            num_m = re.search(r'(\d+[.,]?\d*)', med.replace(',', '.'))
            metros = _to_meters(num_m.group(1), unit_str) if num_m and unit_str else 'A confirmar'

            medicoes.append({
                'item':                     item['item'],
                'id_item':                  item['id'],
                'pagina_slide':             item['pagina_slide_origem'],
                'medida_original':          med,
                'unidade_original':         unit_str,
                'medida_metros':            metros,
                'tipo_medida':              tipo,
                'confiabilidade':           'Média',
                'observacao_confiabilidade': 'Extração automática por análise de texto/tabela',
            })

        # ── Pendências ────────────────────────────────────────────────────────
        pendencias: list[dict] = []
        for item in self.items:
            pend = item.get('pendencia_acao_necessaria', '').strip()
            if not pend:
                continue
            nivel    = item.get('nivel_atencao', 'Médio')
            urgencia = 'Alto' if nivel == 'Alto' else 'Médio'
            pendencias.append({
                'id_item':               item['id'],
                'item':                  item['item'],
                'pagina_slide':          item['pagina_slide_origem'],
                'secao':                 item['secao'],
                'sub_secao_ambiente':    item.get('sub_secao_ambiente', ''),
                'descricao_pendencia':   pend,
                'tipo_pendencia':        'A confirmar',
                'nivel_urgencia':        urgencia,
                'responsavel':           item.get('responsavel_area', 'Não informado'),
                'prazo':                 'Não informado',
                # enrichment used by excel_generator for the Pendências sheet
                'status_arte':           item.get('status_arte', ''),
                'status_compra_locacao': item.get('status_compra_locacao', ''),
                'fonte_informacao':      item.get('fonte_informacao', ''),
            })

        # ── Contagens para resumo ─────────────────────────────────────────────
        def _count_field(field: str) -> dict[str, int]:
            counts: dict[str, int] = {}
            for item in self.items:
                v = item.get(field, '') or 'Não informado'
                counts[v] = counts.get(v, 0) + 1
            return counts

        sec_counts  = {s: len(d['ids']) for s, d in sec_map.items()}
        forn_counts = _count_field('fornecedor')
        status_pool = {}
        for item in self.items:
            sa = item.get('status_arte', '') or item.get('status_compra_locacao', '') or 'Não informado'
            status_pool[sa] = status_pool.get(sa, 0) + 1

        unit_name = 'páginas' if self.doc_type == 'pdf' else 'slides'
        obs_geral = (
            f"Análise automática local. "
            f"{len(self.items)} itens extraídos de {total_units} {unit_name}. "
            f"{len(pendencias)} pendências identificadas. "
            f"{len(medicoes)} medições encontradas."
        )

        return {
            'resumo': {
                'cliente':              self._meta.get('client', 'Não informado'),
                'projeto':              self._meta.get('project', 'Não informado'),
                'total_paginas_slides': total_units,
                'total_itens':          len(self.items),
                'secoes':               [{'nome': s, 'quantidade_itens': c} for s, c in sec_counts.items()],
                'fornecedores':         [{'nome': f, 'quantidade_itens': c} for f, c in forn_counts.items()],
                'status_summary':       [{'status': s, 'quantidade': c} for s, c in status_pool.items()],
                'observacoes_gerais':   obs_geral,
            },
            'itens_detalhados': self.items,
            'por_secao':        por_secao,
            'medicoes':         medicoes,
            'pendencias':       pendencias,
            'legenda_status':   [],
        }
