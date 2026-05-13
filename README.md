# Decupagem Interna

Ferramenta de decupagem técnica de projetos para orçamento, produção e conferência.

Analisa arquivos **PDF** e **PPTX** e gera uma planilha Excel formatada com:

| Aba | Conteúdo |
|-----|----------|
| **Resumo** | Cliente, projeto, totais, distribuição por seção/fornecedor/status |
| **Itens Detalhados** | Todos os itens linha por linha (28 colunas) |
| **Por Seção** | Agrupamento por grandes blocos do projeto |
| **Medições** | Medidas com unidade, conversão em metros e confiabilidade |
| **Pendências** | Ações, aprovações e definições pendentes com nível de urgência |
| **Legenda e Status** | Cores, fornecedores e convenções do projeto |

## Requisitos

- Python 3.11+
- Chave de API Anthropic (Claude)

## Instalação

```bash
pip install -r requirements.txt
```

Copie `.env.example` para `.env` e preencha sua chave:

```bash
cp .env.example .env
# edite .env com ANTHROPIC_API_KEY=sk-ant-...
```

## Execução

```bash
streamlit run app.py
```

Acesse `http://localhost:8501` no navegador.

## Estrutura

```
├── app.py                         # Interface web Streamlit
├── src/
│   ├── extractors/
│   │   ├── pdf_extractor.py       # Extração de PDFs (pdfplumber + PyMuPDF)
│   │   └── pptx_extractor.py      # Extração de PPTX (python-pptx)
│   ├── analyzer/
│   │   └── claude_analyzer.py     # Análise via Claude API
│   └── generator/
│       └: excel_generator.py      # Geração do Excel formatado (openpyxl)
├── requirements.txt
└── .env.example
```
