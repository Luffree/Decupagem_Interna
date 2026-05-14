"""
PDF content extractor using PyMuPDF only (no pdfplumber/pdfminer dependency).
Extracts text, tables, annotations, image metadata and links per page.
"""
import os
import tempfile
import warnings
from dataclasses import dataclass, field

# Suppress the pymupdf_layout suggestion warning
warnings.filterwarnings("ignore", message=".*pymupdf_layout.*")


@dataclass
class PageContent:
    page_number: int
    text: str
    tables: list
    annotations: list
    image_descriptions: list
    hyperlinks: list
    raw_words: list


@dataclass
class PDFContent:
    filename: str
    total_pages: int
    pages: list = field(default_factory=list)
    metadata: dict = field(default_factory=dict)


def _extract_tables_from_page(page) -> list:
    """Use PyMuPDF's built-in table finder."""
    tables = []
    try:
        tab_finder = page.find_tables()
        for tab in tab_finder.tables:
            rows = []
            for row in tab.extract():
                cleaned = [str(cell).strip() if cell is not None else "" for cell in row]
                rows.append(cleaned)
            if rows:
                tables.append(rows)
    except Exception:
        pass
    return tables


def _extract_page(fitz_doc, fitz_page) -> PageContent:
    import fitz as _fitz
    page_num = fitz_page.number + 1

    # Text (layout-preserving)
    text = fitz_page.get_text("text") or ""

    # Tables via find_tables
    tables = _extract_tables_from_page(fitz_page)

    # Annotations
    annots = []
    for annot in fitz_page.annots():
        info = annot.info or {}
        content = info.get("content", "").strip()
        author  = info.get("title", "").strip()
        if content or author:
            annots.append({
                "type":    annot.type[1] if annot.type else "unknown",
                "content": content,
                "author":  author,
                "rect":    list(annot.rect),
            })

    # Images (metadata only)
    images = []
    for img in fitz_page.get_images(full=True):
        xref = img[0]
        try:
            base = fitz_doc.extract_image(xref)
            images.append(
                f"Imagem ({base.get('width','?')}x{base.get('height','?')} px, "
                f"tipo: {base.get('ext','?')})"
            )
        except Exception:
            images.append("Imagem (metadados indisponíveis)")

    # Hyperlinks
    links = [lk.get("uri", "") for lk in fitz_page.get_links() if lk.get("uri")]

    # Raw words with positions
    raw_words = [
        {"text": w[4], "x0": w[0], "top": w[1]}
        for w in fitz_page.get_text("words")
    ]

    return PageContent(
        page_number=page_num,
        text=text,
        tables=tables,
        annotations=annots,
        image_descriptions=images,
        hyperlinks=links,
        raw_words=raw_words,
    )


def extract_pdf(file_path: str) -> PDFContent:
    import fitz
    doc = fitz.open(file_path)
    content = PDFContent(
        filename=file_path,
        total_pages=len(doc),
        metadata=dict(doc.metadata) if doc.metadata else {},
    )
    for page in doc:
        content.pages.append(_extract_page(doc, page))
    doc.close()
    return content


def extract_pdf_bytes(file_bytes: bytes, filename: str) -> PDFContent:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        result = extract_pdf(tmp_path)
        result.filename = filename
        return result
    finally:
        os.unlink(tmp_path)


def pdf_content_to_dict(content: PDFContent) -> dict:
    return {
        "filename":    content.filename,
        "total_pages": content.total_pages,
        "metadata":    content.metadata,
        "pages": [
            {
                "page_number":       p.page_number,
                "text":              p.text,
                "tables":            p.tables,
                "annotations":       p.annotations,
                "image_descriptions": p.image_descriptions,
                "hyperlinks":        p.hyperlinks,
            }
            for p in content.pages
        ],
    }
