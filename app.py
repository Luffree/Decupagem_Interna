"""
Decupagem Interna — Interface web Streamlit (sem API externa).
Processa PDF/PPTX e gera planilha Excel de decupagem técnica.
"""
import io
import time
import datetime
import traceback
from pathlib import Path

import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Decupagem Interna",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { padding-top: 1.8rem; }

    .app-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 1.8rem 2.2rem;
        border-radius: 12px;
        margin-bottom: 1.8rem;
        color: white;
    }
    .app-header h1 { margin: 0; font-size: 1.9rem; font-weight: 700; }
    .app-header p  { margin: 0.4rem 0 0; opacity: 0.88; font-size: 0.97rem; }

    .info-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.1rem 1.4rem;
        margin-bottom: 0.9rem;
    }
    .info-card h4 { color: #1F3864; margin: 0 0 0.4rem; font-size: 0.93rem; font-weight: 700; }
    .info-card p  { color: #475569; margin: 0; font-size: 0.84rem; line-height: 1.55; }

    .step-badge {
        display: inline-block;
        background: #1F3864; color: white;
        border-radius: 50%; width: 26px; height: 26px;
        line-height: 26px; text-align: center;
        font-weight: 700; font-size: 0.88rem; margin-right: 0.45rem;
    }

    .badge-alto    { background:#FF6B6B; color:white;  padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
    .badge-medio   { background:#FFD93D; color:#333;   padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
    .badge-baixo   { background:#6BCB77; color:white;  padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
    .badge-critico { background:#C0392B; color:white;  padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }

    .summary-table { width:100%; border-collapse:collapse; font-size:0.85rem; }
    .summary-table th { background:#1F3864; color:white; padding:7px 11px; text-align:left; }
    .summary-table td { padding:6px 11px; border-bottom:1px solid #e2e8f0; }
    .summary-table tr:nth-child(even) td { background:#f8fafc; }

    div[data-testid="stMetric"] { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:0.6rem 0.9rem; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📋 Decupagem Interna")
    st.markdown("*Análise local — sem API*")
    st.markdown("---")

    st.markdown("### 📊 Abas geradas")
    for icon, label, desc in [
        ("📋", "Resumo",           "Cliente, projeto, totais"),
        ("📝", "Itens Detalhados", "28 colunas por item"),
        ("📁", "Por Seção",        "Blocos do projeto"),
        ("📐", "Medições",         "Medidas + conversão"),
        ("⚠️", "Pendências",       "Ações e aprovações"),
        ("🎨", "Legenda e Status", "Cores e convenções"),
    ]:
        st.markdown(f"**{icon} {label}** — {desc}")

    st.markdown("---")
    st.markdown("### 🎨 Nível de Atenção")
    st.markdown("""
    <span class="badge-baixo">Baixo</span> Sem pendências<br>
    <span class="badge-medio">Médio</span> Pendências controláveis<br>
    <span class="badge-alto">Alto</span> Item crítico<br>
    <span class="badge-critico">Crítico</span> Bloqueante / urgente
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📌 O que é detectado")
    st.markdown("""
    - Textos, tabelas, anotações
    - Medidas (L×A×P, cm, mm, m²)
    - Materiais (MDF, aço, acrílico…)
    - Cores e acabamentos
    - Categorias (cenografia, mobiliário…)
    - Status de arte e compra
    - Pendências e itens críticos
    - Notas e formas de slides
    """)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📋 Decupagem Interna</h1>
    <p>Extração e organização técnica de projetos — orçamento, produção e conferência · Análise 100% local</p>
</div>
""", unsafe_allow_html=True)


# ── Upload ────────────────────────────────────────────────────────────────────
col_up, col_info = st.columns([3, 1])

with col_up:
    st.markdown('<span class="step-badge">1</span> **Carregue o arquivo do projeto**', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Arraste ou clique para selecionar um PDF ou PPTX",
        type=["pdf", "pptx", "ppt"],
        label_visibility="collapsed",
    )

with col_info:
    st.markdown("""
    <div class="info-card">
        <h4>Formatos aceitos</h4>
        <p><b>.pdf</b> — Documentos PDF<br><b>.pptx / .ppt</b> — PowerPoint</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _run_extraction(file_bytes: bytes, filename: str, file_ext: str, status_ph, prog_ph):
    """Run extraction + local analysis. Returns (analyzed_data, doc_type, total_units, unit_label)."""
    logs: list[str] = []

    def cb(msg: str):
        logs.append(f"• {msg}")
        status_ph.markdown("**Status:** " + msg)

    cb("Extraindo conteúdo do arquivo...")
    t0 = time.time()

    if file_ext == ".pdf":
        from src.extractors.pdf_extractor import extract_pdf_bytes, pdf_content_to_dict
        raw = extract_pdf_bytes(file_bytes, filename)
        doc_dict   = pdf_content_to_dict(raw)
        doc_type   = "pdf"
        total_units = raw.total_pages
        unit_label  = "páginas"
    else:
        from src.extractors.pptx_extractor import extract_pptx_bytes, pptx_content_to_dict
        raw = extract_pptx_bytes(file_bytes, filename)
        doc_dict   = pptx_content_to_dict(raw)
        doc_type   = "pptx"
        total_units = raw.total_slides
        unit_label  = "slides"

    prog_ph.progress(25, text="Extração concluída")
    cb(f"Extração concluída: {total_units} {unit_label} em {time.time()-t0:.1f}s")

    cb("Analisando conteúdo com motor local...")
    from src.analyzer.local_analyzer import analyze_document

    def on_prog(msg: str):
        logs.append(f"  ↳ {msg}")
        status_ph.markdown("**Status:** " + msg)

    t1 = time.time()
    prog_ph.progress(35, text="Análise em andamento...")
    analyzed = analyze_document(doc_dict, doc_type, progress_callback=on_prog)
    prog_ph.progress(80, text="Análise concluída")
    cb(f"Análise concluída em {time.time()-t1:.1f}s — "
       f"{len(analyzed.get('itens_detalhados', []))} itens, "
       f"{len(analyzed.get('pendencias', []))} pendências")

    cb("Gerando planilha Excel formatada...")
    from src.generator.excel_generator import generate_excel
    excel_bytes = generate_excel(analyzed, filename)
    prog_ph.progress(100, text="Planilha pronta!")
    cb("Planilha pronta para download!")

    return analyzed, doc_type, total_units, unit_label, excel_bytes, logs


# ── Main flow ─────────────────────────────────────────────────────────────────

if uploaded is not None:
    file_ext   = Path(uploaded.name).suffix.lower()
    file_bytes = uploaded.read()

    # File info metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("Arquivo",  uploaded.name[:35] + ("…" if len(uploaded.name) > 35 else ""))
    m2.metric("Tamanho",  f"{len(file_bytes)/1024:.1f} KB")
    m3.metric("Formato",  file_ext.upper())

    st.markdown("---")
    st.markdown('<span class="step-badge">2</span> **Processar arquivo**', unsafe_allow_html=True)

    if st.button("🚀 Iniciar Decupagem Técnica", type="primary", use_container_width=True):
        prog_ph   = st.progress(0, text="Iniciando…")
        status_ph = st.empty()
        log_exp   = st.expander("📜 Log de processamento", expanded=True)

        try:
            analyzed, doc_type, total_units, unit_label, excel_bytes, logs = \
                _run_extraction(file_bytes, uploaded.name, file_ext, status_ph, prog_ph)

            with log_exp:
                st.markdown("\n".join(logs))

            prog_ph.empty()
            status_ph.empty()

            # ── Results ──────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown('<span class="step-badge">3</span> **Resultado da Decupagem**', unsafe_allow_html=True)

            resumo   = analyzed.get('resumo', {})
            n_itens  = len(analyzed.get('itens_detalhados', []))
            n_pend   = len(analyzed.get('pendencias', []))
            n_med    = len(analyzed.get('medicoes', []))
            n_secoes = len(resumo.get('secoes', []))

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Itens",             n_itens)
            c2.metric("Pendências",        n_pend)
            c3.metric("Medições",          n_med)
            c4.metric("Seções",            n_secoes)
            c5.metric(unit_label.capitalize(), total_units)

            # Sections table
            secoes = resumo.get('secoes', [])
            if secoes:
                st.markdown("#### 📁 Itens por Seção")
                rows_html = "".join(
                    f"<tr><td>{s['nome']}</td>"
                    f"<td style='text-align:center;font-weight:600'>{s['quantidade_itens']}</td></tr>"
                    for s in secoes
                )
                st.markdown(
                    f'<table class="summary-table"><thead><tr><th>Seção</th><th>Qtd.</th></tr></thead>'
                    f'<tbody>{rows_html}</tbody></table>',
                    unsafe_allow_html=True,
                )
                st.markdown("")

            # Top category breakdown
            cats: dict[str, int] = {}
            for item in analyzed.get('itens_detalhados', []):
                cat = item.get('categoria', 'A confirmar') or 'A confirmar'
                cats[cat] = cats.get(cat, 0) + 1
            if cats:
                st.markdown("#### 🗂️ Itens por Categoria")
                sorted_cats = sorted(cats.items(), key=lambda x: -x[1])
                rows_html = "".join(
                    f"<tr><td>{cat}</td>"
                    f"<td style='text-align:center;font-weight:600'>{count}</td></tr>"
                    for cat, count in sorted_cats
                )
                st.markdown(
                    f'<table class="summary-table"><thead><tr><th>Categoria</th><th>Qtd.</th></tr></thead>'
                    f'<tbody>{rows_html}</tbody></table>',
                    unsafe_allow_html=True,
                )
                st.markdown("")

            # Critical / high pendências
            pendencias = analyzed.get('pendencias', [])
            criticas = [p for p in pendencias
                        if str(p.get('nivel_urgencia', '')).lower() in ('alto', 'crítico', 'critico')]
            if criticas:
                with st.expander(f"⚠️ Pendências Altas / Críticas ({len(criticas)})"):
                    for p in criticas[:15]:
                        urg   = str(p.get('nivel_urgencia', '')).strip()
                        badge = 'badge-critico' if urg.lower() in ('crítico', 'critico') else 'badge-alto'
                        st.markdown(
                            f'<span class="{badge}">{urg}</span> '
                            f"**{p.get('item', 'Item')}** "
                            f"({unit_label} {p.get('pagina_slide', '?')}): "
                            f"{p.get('descricao_pendencia', '')}",
                            unsafe_allow_html=True,
                        )

            # Obs gerais
            obs = resumo.get('observacoes_gerais', '')
            if obs:
                with st.expander("📝 Observações gerais"):
                    st.write(obs)

            # ── Download ─────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown('<span class="step-badge">4</span> **Download da Planilha**', unsafe_allow_html=True)

            ts          = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            stem        = Path(uploaded.name).stem
            output_name = f"decupagem_{stem}_{ts}.xlsx"

            st.download_button(
                label="⬇️  Baixar Planilha Excel (.xlsx)",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            prog_ph.empty()
            st.error(f"❌ Erro durante o processamento: {e}")
            with st.expander("Detalhes técnicos do erro"):
                st.code(traceback.format_exc())

else:
    # Landing info when no file is loaded
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("""
        <div class="info-card">
            <h4>🚀 Como funciona</h4>
            <p>
                <b>1.</b> Carregue um PDF ou PPTX de projeto técnico<br>
                <b>2.</b> O sistema extrai textos, tabelas, anotações, formas e notas<br>
                <b>3.</b> O motor local analisa e categoriza cada item individualmente<br>
                <b>4.</b> Uma planilha Excel completa é gerada com 6 abas temáticas<br>
                <b>5.</b> Baixe o arquivo pronto para orçamento e conferência
            </p>
        </div>
        """, unsafe_allow_html=True)

    with col_b:
        st.markdown("""
        <div class="info-card">
            <h4>📌 O que é extraído automaticamente</h4>
            <p>
                Medidas (L×A×P, cm, mm, m², metro linear) ·
                Materiais (MDF, aço, acrílico, lona…) ·
                Cores e acabamentos ·
                Categorias (cenografia, mobiliário, CV, elétrica…) ·
                Status de arte e compra/locação ·
                Pendências com nível de urgência ·
                Tabelas, anotações, notas de slide, formas com cor
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-card">
        <h4>💡 Dicas para melhor resultado</h4>
        <p>
            • PDFs com texto selecionável produzem resultados melhores que imagens escaneadas<br>
            • Tabelas bem estruturadas no documento geram itens mais detalhados<br>
            • Títulos de slides no PPTX são usados como nomes de seção<br>
            • Medidas no formato <b>L × A</b> ou <b>L × A × P</b> são detectadas automaticamente<br>
            • Palavras como "pendente", "a confirmar" e "aguardando" são identificadas como pendências
        </p>
    </div>
    """, unsafe_allow_html=True)


st.markdown("---")
st.markdown(
    '<p style="text-align:center;color:#94a3b8;font-size:0.78rem;">'
    'Decupagem Interna · Análise técnica local para produção e orçamento'
    '</p>',
    unsafe_allow_html=True,
)
