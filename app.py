"""
Decupagem Interna — Streamlit web interface.
Processes PDF/PPTX files and generates a formatted Excel decupagem spreadsheet.
"""
import os
import time
import tempfile
import datetime
import streamlit as st
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Decupagem Interna",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .main .block-container { padding-top: 2rem; }

    /* Header */
    .app-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        color: white;
    }
    .app-header h1 { margin: 0; font-size: 2rem; font-weight: 700; }
    .app-header p { margin: 0.5rem 0 0; opacity: 0.85; font-size: 1rem; }

    /* Cards */
    .info-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1rem;
    }
    .info-card h4 { color: #1F3864; margin: 0 0 0.5rem; font-size: 0.95rem; }
    .info-card p { color: #475569; margin: 0; font-size: 0.85rem; line-height: 1.5; }

    /* Step badges */
    .step-badge {
        display: inline-block;
        background: #1F3864;
        color: white;
        border-radius: 50%;
        width: 28px;
        height: 28px;
        line-height: 28px;
        text-align: center;
        font-weight: 700;
        font-size: 0.9rem;
        margin-right: 0.5rem;
    }

    /* Status badges */
    .badge-alto    { background:#FF6B6B; color:white; padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
    .badge-medio   { background:#FFD93D; color:#333; padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
    .badge-baixo   { background:#6BCB77; color:white; padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
    .badge-critico { background:#C0392B; color:white; padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }

    /* Divider */
    hr { border: none; border-top: 2px solid #e2e8f0; margin: 1.5rem 0; }

    /* Upload area */
    .uploadedFile { border: 2px solid #2E75B6 !important; }

    /* Success box */
    .success-box {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
    }

    /* Summary table */
    .summary-table { width: 100%; border-collapse: collapse; }
    .summary-table th {
        background: #1F3864; color: white;
        padding: 8px 12px; text-align: left;
        font-size: 0.85rem;
    }
    .summary-table td {
        padding: 7px 12px;
        border-bottom: 1px solid #e2e8f0;
        font-size: 0.85rem;
    }
    .summary-table tr:nth-child(even) td { background: #f8fafc; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuração")
    st.markdown("---")

    api_key = st.text_input(
        "Chave API Anthropic (Claude)",
        value=os.environ.get("ANTHROPIC_API_KEY", ""),
        type="password",
        help="Sua chave da API Anthropic. Pode ser definida também via variável de ambiente ANTHROPIC_API_KEY.",
    )

    st.markdown("---")
    st.markdown("### 📊 Sobre as Abas Geradas")
    tabs_info = [
        ("📋 Resumo",           "Visão geral do projeto, cliente, totais e distribuição"),
        ("📝 Itens Detalhados", "Todos os itens linha por linha com 28 colunas"),
        ("📁 Por Seção",        "Agrupamento por grandes blocos do projeto"),
        ("📐 Medições",         "Todas as medidas com confiabilidade"),
        ("⚠️ Pendências",       "Ações, aprovações e definições pendentes"),
        ("🎨 Legenda e Status", "Cores, fornecedores e convenções do projeto"),
    ]
    for icon_tab, desc in tabs_info:
        st.markdown(f"**{icon_tab}**  \n{desc}")

    st.markdown("---")
    st.markdown("### 🎨 Código de Cores")
    st.markdown("""
    <span class="badge-baixo">Baixo</span> Sem pendências críticas
    <span class="badge-medio">Médio</span> Pendências controladas
    <span class="badge-alto">Alto</span> Item crítico
    <span class="badge-critico">Crítico</span> Bloqueante / urgente
    """, unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📋 Decupagem Interna</h1>
    <p>Extração e organização técnica de projetos para orçamento, produção e conferência</p>
</div>
""", unsafe_allow_html=True)


# ── Upload section ────────────────────────────────────────────────────────────
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown('<span class="step-badge">1</span> **Carregue o arquivo do projeto**', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Arraste ou clique para selecionar",
        type=["pdf", "pptx", "ppt"],
        help="Formatos suportados: PDF, PPTX, PPT",
        label_visibility="collapsed",
    )

with col2:
    st.markdown('<div class="info-card"><h4>ℹ️ Formatos aceitos</h4><p><b>.pdf</b> — Documentos PDF (texto, tabelas, anotações, imagens)<br><b>.pptx</b> — Apresentações PowerPoint (slides, notas, formas, tabelas)</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ── Process ───────────────────────────────────────────────────────────────────
if uploaded_file is not None:
    file_ext = Path(uploaded_file.name).suffix.lower()
    file_bytes = uploaded_file.read()

    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.metric("Arquivo", uploaded_file.name[:30] + ("..." if len(uploaded_file.name) > 30 else ""))
    with col_info2:
        st.metric("Tamanho", f"{len(file_bytes) / 1024:.1f} KB")
    with col_info3:
        st.metric("Formato", file_ext.upper())

    st.markdown("---")
    st.markdown('<span class="step-badge">2</span> **Processar arquivo**', unsafe_allow_html=True)

    if not api_key:
        st.warning("⚠️ Insira sua chave API Anthropic na barra lateral antes de processar.")
    else:
        if st.button("🚀 Iniciar Decupagem Técnica", type="primary", use_container_width=True):
            try:
                progress_bar = st.progress(0, text="Iniciando...")
                status_text  = st.empty()
                log_area     = st.expander("📜 Log de processamento", expanded=True)
                log_lines    = []

                def update_progress(msg: str, pct: int = None):
                    log_lines.append(f"• {msg}")
                    with log_area:
                        st.markdown("\n".join(log_lines))
                    if pct is not None:
                        progress_bar.progress(pct, text=msg)
                    status_text.markdown(f"**Status:** {msg}")

                # ── Step 1: Extract ──────────────────────────────────────────
                update_progress("Extraindo conteúdo do arquivo...", 10)
                t0 = time.time()

                if file_ext == ".pdf":
                    from src.extractors.pdf_extractor import extract_pdf_bytes, pdf_content_to_dict
                    raw_content = extract_pdf_bytes(file_bytes, uploaded_file.name)
                    doc_dict    = pdf_content_to_dict(raw_content)
                    doc_type    = "pdf"
                    total_units = raw_content.total_pages
                    unit_label  = "páginas"
                else:
                    from src.extractors.pptx_extractor import extract_pptx_bytes, pptx_content_to_dict
                    raw_content = extract_pptx_bytes(file_bytes, uploaded_file.name)
                    doc_dict    = pptx_content_to_dict(raw_content)
                    doc_type    = "pptx"
                    total_units = raw_content.total_slides
                    unit_label  = "slides"

                elapsed = time.time() - t0
                update_progress(f"Extração concluída: {total_units} {unit_label} em {elapsed:.1f}s", 30)

                # ── Step 2: Analyze ──────────────────────────────────────────
                update_progress("Enviando para análise pelo Claude (pode levar alguns minutos)...", 40)
                from src.analyzer.claude_analyzer import analyze_document

                def on_progress(msg: str):
                    update_progress(msg)

                t1 = time.time()
                analyzed_data = analyze_document(
                    doc_dict,
                    doc_type,
                    api_key,
                    progress_callback=on_progress,
                )
                elapsed2 = time.time() - t1
                n_itens = len(analyzed_data.get("itens_detalhados", []))
                n_pend  = len(analyzed_data.get("pendencias", []))
                update_progress(f"Análise concluída em {elapsed2:.1f}s — {n_itens} itens, {n_pend} pendências", 75)

                # ── Step 3: Generate Excel ───────────────────────────────────
                update_progress("Gerando planilha Excel formatada...", 85)
                from src.generator.excel_generator import generate_excel

                excel_bytes = generate_excel(analyzed_data, uploaded_file.name)
                update_progress("Planilha Excel pronta!", 100)
                progress_bar.empty()
                status_text.empty()

                # ── Results ──────────────────────────────────────────────────
                st.markdown("---")
                st.markdown('<span class="step-badge">3</span> **Resultado**', unsafe_allow_html=True)

                resumo = analyzed_data.get("resumo", {})

                res_col1, res_col2, res_col3, res_col4 = st.columns(4)
                res_col1.metric("Total de Itens",       resumo.get("total_itens", n_itens))
                res_col2.metric("Pendências",            n_pend)
                res_col3.metric("Medições",              len(analyzed_data.get("medicoes", [])))
                res_col4.metric("Seções identificadas",  len(resumo.get("secoes", [])))

                st.markdown("---")

                # Summary table
                secoes = resumo.get("secoes", [])
                if secoes:
                    st.markdown("#### 📁 Itens por Seção")
                    rows_html = "".join(
                        f"<tr><td>{s.get('nome', '')}</td><td style='text-align:center'>{s.get('quantidade_itens', 0)}</td></tr>"
                        for s in secoes
                    )
                    st.markdown(
                        f'<table class="summary-table"><thead><tr><th>Seção</th><th>Qtd.</th></tr></thead>'
                        f'<tbody>{rows_html}</tbody></table>',
                        unsafe_allow_html=True,
                    )
                    st.markdown("")

                # Download button
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = Path(uploaded_file.name).stem
                output_name = f"decupagem_{stem}_{ts}.xlsx"

                st.download_button(
                    label="⬇️  Baixar Planilha Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

                # Obs gerais
                obs = resumo.get("observacoes_gerais", "")
                if obs and obs not in ("Não informado", ""):
                    with st.expander("📝 Observações gerais do projeto"):
                        st.write(obs)

                # Top pendências
                pendencias = analyzed_data.get("pendencias", [])
                criticas = [p for p in pendencias if str(p.get("nivel_urgencia", "")).lower() in ("crítico", "critico", "alto")]
                if criticas:
                    with st.expander(f"⚠️ Pendências críticas / altas ({len(criticas)})"):
                        for p in criticas[:10]:
                            urg = str(p.get("nivel_urgencia", "")).strip()
                            badge_class = "badge-critico" if urg.lower() in ("crítico", "critico") else "badge-alto"
                            st.markdown(
                                f'<span class="{badge_class}">{urg}</span> '
                                f'**{p.get("item", "Item")}** (Slide/Pág. {p.get("pagina_slide", "?")}): '
                                f'{p.get("descricao_pendencia", "")}',
                                unsafe_allow_html=True,
                            )

            except Exception as e:
                st.error(f"❌ Erro durante o processamento: {e}")
                with st.expander("Detalhes do erro"):
                    import traceback
                    st.code(traceback.format_exc())

else:
    # Placeholder when no file uploaded
    st.markdown("""
    <div class="info-card">
        <h4>🚀 Como funciona</h4>
        <p>
            <b>1.</b> Carregue um arquivo PDF ou PPTX de projeto técnico<br>
            <b>2.</b> O sistema extrai todo o conteúdo: textos, tabelas, anotações, medidas, imagens, notas e formas<br>
            <b>3.</b> O Claude analisa e categoriza cada item técnico individualmente<br>
            <b>4.</b> Uma planilha Excel formatada é gerada com 6 abas: Resumo, Itens Detalhados, Por Seção, Medições, Pendências e Legenda<br>
            <b>5.</b> Baixe o arquivo pronto para orçamento, produção e conferência
        </p>
    </div>

    <div class="info-card">
        <h4>📌 O que é extraído</h4>
        <p>
            Textos corridos · Tabelas · Legendas · Cotas · Medidas em desenhos ·
            Anotações visuais · Chamadas com setas · Observações técnicas ·
            Listas de materiais · Indicações de fornecedor · Status e pendências ·
            Notas de slide · Formas com cores significativas · Hiperlinks
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.markdown(
    '<p style="text-align:center; color:#94a3b8; font-size:0.8rem;">'
    'Decupagem Interna · Powered by Claude (Anthropic) · Análise técnica para produção e orçamento'
    '</p>',
    unsafe_allow_html=True,
)
