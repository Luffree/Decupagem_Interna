# Decupagem Interna

Ferramenta de decupagem técnica de projetos para orçamento, produção e conferência.
Processa arquivos **PDF** e **PPTX** localmente — sem API externa, sem custos de uso.

## O que é gerado

| Aba | Conteúdo |
|-----|----------|
| **Resumo** | Cliente, projeto, totais, distribuição por seção / fornecedor / status |
| **Itens Detalhados** | Todos os itens linha por linha (28 colunas) |
| **Por Seção** | Agrupamento por grandes blocos do projeto |
| **Medições** | Medidas com unidade, conversão em metros e confiabilidade |
| **Pendências** | Ações, aprovações e definições pendentes com nível de urgência |
| **Legenda e Status** | Cores e convenções do projeto |

## O que é detectado automaticamente

- Medidas (L×A×P, cm, mm, m², metro linear, diâmetro, espessura)
- Materiais (MDF, aço, acrílico, lona, metalon, vidro…)
- Cores e acabamentos (fosco, brilhante, escovado…)
- Categorias (cenografia, mobiliário, comunicação visual, elétrica…)
- Status de arte e compra/locação
- Pendências com nível de urgência (Baixo / Médio / Alto)
- Tabelas, anotações, notas de slide e formas com cor

## Requisitos

- Python 3.11+

## Instalação

```bash
pip install -r requirements.txt
```

## Execução (interface web)

```bash
streamlit run app.py
```

Acesse `http://localhost:8501` no navegador.

## Estrutura

```
├── app.py                          # Interface Streamlit (sem API)
├── src/
│   ├── extractors/
│   │   ├── pdf_extractor.py        # Extração PDF (pdfplumber + PyMuPDF)
│   │   └── pptx_extractor.py       # Extração PPTX (python-pptx)
│   ├── analyzer/
│   │   ├── patterns.py             # Padrões, regex e palavras-chave
│   │   └── local_analyzer.py       # Análise local sem API
│   └── generator/
│       └── excel_generator.py      # Geração do Excel formatado (openpyxl)
└── requirements.txt
```
