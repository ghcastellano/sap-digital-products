"""
PDF Parser — Extract text from PDF files using pdfplumber
"""
from pathlib import Path

import pdfplumber


def parse_pdf(file_path: Path) -> dict:
    """
    Parse a PDF and return structured content by page.

    Returns:
        {
            'pages': [{'number': 1, 'text': '...'}],
            'metadata': {'page_count': N}
        }
    """
    pages_data = []
    with pdfplumber.open(str(file_path)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            pages_data.append({
                "number": i,
                "text": text.strip(),
            })

    return {
        "pages": pages_data,
        "metadata": {"page_count": len(pages_data)},
    }
