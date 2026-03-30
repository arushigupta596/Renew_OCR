import os
import base64
import glob
import fitz  # PyMuPDF
from openai import OpenAI


def extract_pdf_text(pdf_path: str) -> str:
    """Extract text from a PDF using PyMuPDF. Fast, no ML needed."""
    doc = fitz.open(pdf_path)
    pages = []
    for page_num, page in enumerate(doc):
        text = page.get_text("text")
        if text.strip():
            pages.append(f"--- Page {page_num + 1} ---\n{text}")
    doc.close()
    return "\n\n".join(pages)


def extract_pdf_with_vision(
    pdf_path: str, api_key: str, model: str, base_url: str
) -> str:
    """
    For scanned/image PDFs with no embedded text,
    convert pages to images and send to a vision LLM.
    """
    doc = fitz.open(pdf_path)
    all_text = []

    client = OpenAI(api_key=api_key, base_url=base_url)

    for page_num in range(len(doc)):
        page = doc[page_num]
        # Render page as image
        pix = page.get_pixmap(dpi=200)
        img_bytes = pix.tobytes("png")
        b64_img = base64.b64encode(img_bytes).decode("utf-8")

        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Extract ALL text from this document image. Preserve the layout, tables, and structure as much as possible. Return only the extracted text, nothing else.",
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{b64_img}"
                            },
                        },
                    ],
                }
            ],
            temperature=0,
            max_tokens=4000,
        )
        page_text = response.choices[0].message.content
        all_text.append(f"--- Page {page_num + 1} ---\n{page_text}")

    doc.close()
    return "\n\n".join(all_text)


def _has_text(pdf_path: str) -> bool:
    """Check if a PDF has extractable text (not just scanned images)."""
    doc = fitz.open(pdf_path)
    total_text = ""
    for page in doc:
        total_text += page.get_text("text")
        if len(total_text.strip()) > 50:
            doc.close()
            return True
    doc.close()
    return False


def extract_all_pdfs(
    folder_path: str,
    api_key: str = "",
    vision_model: str = "",
    base_url: str = "",
    progress_callback=None,
    **kwargs,
) -> dict[str, str]:
    """
    Extract text from all PDFs in a folder.

    Uses fast PyMuPDF text extraction for PDFs with embedded text.
    Falls back to vision LLM for scanned/image-only PDFs.

    Returns a dict mapping filename to extracted text.
    """
    pdf_files = sorted(glob.glob(os.path.join(folder_path, "*.pdf")))
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in {folder_path}")

    results = {}
    total = len(pdf_files)

    for i, pdf_path in enumerate(pdf_files):
        filename = os.path.basename(pdf_path)

        try:
            if _has_text(pdf_path):
                # Fast path: extract embedded text
                text = extract_pdf_text(pdf_path)
                results[filename] = text
            elif api_key:
                # Slow path: use vision LLM for scanned PDFs
                text = extract_pdf_with_vision(pdf_path, api_key, vision_model, base_url)
                results[filename] = text
            else:
                results[filename] = "[OCR ERROR] Scanned PDF detected but no API key for vision OCR"
        except Exception as e:
            results[filename] = f"[OCR ERROR] {str(e)}"

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return results
