"""
PDF content extractor using pdfplumber and PyMuPDF.
Extracts text, tables, annotations, images metadata and visual elements.
"""
import io
import base64
from dataclasses import dataclass, field
from typing import Optional
import pdfplumber
import fitz  # PyMuPDF


@dataclass
class PageContent:
    page_number: int
    text: str
    tables: list[list[list[str]]]
    annotations: list[dict]
    image_descriptions: list[str]
    hyperlinks: list[str]
    raw_words: list[dict]


@dataclass
class PDFContent:
    filename: str
    total_pages: int
    pages: list[PageContent] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)


def extract_pdf(file_path: str) -> PDFContent:
    content = PDFContent(
        filename=file_path,
        total_pages=0,
        pages=[],
        metadata={},
    )

    # Extract with PyMuPDF for annotations, links, images
    fitz_doc = fitz.open(file_path)
    content.total_pages = len(fitz_doc)
    content.metadata = dict(fitz_doc.metadata) if fitz_doc.metadata else {}

    fitz_pages_data: dict[int, dict] = {}
    for page_index, fitz_page in enumerate(fitz_doc):
        page_num = page_index + 1
        annots = []
        for annot in fitz_page.annots():
            annot_info = {
                "type": annot.type[1] if annot.type else "unknown",
                "content": annot.info.get("content", ""),
                "author": annot.info.get("title", ""),
                "rect": list(annot.rect),
            }
            if annot_info["content"] or annot_info["author"]:
                annots.append(annot_info)

        links = []
        for link in fitz_page.get_links():
            uri = link.get("uri", "")
            if uri:
                links.append(uri)

        images = []
        for img in fitz_page.get_images(full=True):
            xref = img[0]
            try:
                base_image = fitz_doc.extract_image(xref)
                images.append(f"Imagem ({base_image.get('width', '?')}x{base_image.get('height', '?')} px, tipo: {base_image.get('ext', '?')})")
            except Exception:
                images.append("Imagem (metadados indisponíveis)")

        fitz_pages_data[page_num] = {
            "annotations": annots,
            "links": links,
            "images": images,
        }

    fitz_doc.close()

    # Extract text and tables with pdfplumber
    with pdfplumber.open(file_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            page_num = page_index + 1
            text = page.extract_text(layout=True) or ""

            tables = []
            for table in page.extract_tables():
                cleaned = []
                for row in table:
                    cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                    cleaned.append(cleaned_row)
                if cleaned:
                    tables.append(cleaned)

            words = page.extract_words() or []
            raw_words = [
                {"text": w["text"], "x0": w["x0"], "top": w["top"]}
                for w in words
            ]

            fitz_data = fitz_pages_data.get(page_num, {})

            content.pages.append(PageContent(
                page_number=page_num,
                text=text,
                tables=tables,
                annotations=fitz_data.get("annotations", []),
                image_descriptions=fitz_data.get("images", []),
                hyperlinks=fitz_data.get("links", []),
                raw_words=raw_words,
            ))

    return content


def extract_pdf_bytes(file_bytes: bytes, filename: str) -> PDFContent:
    """Extract content from PDF given as bytes."""
    import tempfile, os
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
    """Convert PDFContent to a JSON-serializable dict for the analyzer."""
    return {
        "filename": content.filename,
        "total_pages": content.total_pages,
        "metadata": content.metadata,
        "pages": [
            {
                "page_number": p.page_number,
                "text": p.text,
                "tables": p.tables,
                "annotations": p.annotations,
                "image_descriptions": p.image_descriptions,
                "hyperlinks": p.hyperlinks,
            }
            for p in content.pages
        ],
    }
