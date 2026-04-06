import os
import base64
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed
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


def _extract_pdf_page_texts(pdf_path: str) -> list[str]:
    """Extract plain text for each PDF page."""
    doc = fitz.open(pdf_path)
    pages: list[str] = []
    for page in doc:
        pages.append(page.get_text("text"))
    doc.close()
    return pages


def _render_page_b64(pdf_path: str, page_num: int, dpi: int = 150) -> str:
    """Render a single PDF page to a base64-encoded PNG."""
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    pix = page.get_pixmap(dpi=dpi)
    img_bytes = pix.tobytes("png")
    doc.close()
    return base64.b64encode(img_bytes).decode("utf-8")


def extract_pdf_with_vision(
    pdf_path: str, api_key: str, model: str, base_url: str, max_workers: int = 4
) -> str:
    """
    Convert PDF pages to images and send to a vision LLM in parallel.
    Pages are processed concurrently (up to max_workers at a time).
    """
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    doc.close()

    client = OpenAI(api_key=api_key, base_url=base_url)

    def _process_page(page_num: int) -> tuple[int, str]:
        b64_img = _render_page_b64(pdf_path, page_num)
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
                            "image_url": {"url": f"data:image/png;base64,{b64_img}"},
                        },
                    ],
                }
            ],
            temperature=0,
            max_tokens=4000,
        )
        return page_num, response.choices[0].message.content

    results = {}
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_process_page, i): i for i in range(num_pages)}
        for future in as_completed(futures):
            page_num, text = future.result()
            results[page_num] = text

    return "\n\n".join(
        f"--- Page {i + 1} ---\n{results[i]}" for i in range(num_pages)
    )


def extract_pdf_hybrid(
    pdf_path: str, api_key: str, model: str, base_url: str, max_workers: int = 4
) -> str:
    """Extract text from all pages, using vision only for pages without text."""
    page_texts = _extract_pdf_page_texts(pdf_path)
    missing_pages = [i for i, text in enumerate(page_texts) if not text.strip()]
    if not missing_pages:
        return "\n\n".join(
            f"--- Page {i + 1} ---\n{text}" for i, text in enumerate(page_texts) if text.strip()
        )

    client = OpenAI(api_key=api_key, base_url=base_url)

    def _process_page(page_num: int) -> tuple[int, str]:
        b64_img = _render_page_b64(pdf_path, page_num)
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
                            "image_url": {"url": f"data:image/png;base64,{b64_img}"},
                        },
                    ],
                }
            ],
            temperature=0,
            max_tokens=4000,
        )
        return page_num, response.choices[0].message.content

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_process_page, i): i for i in missing_pages}
        for future in as_completed(futures):
            page_num, text = future.result()
            page_texts[page_num] = text

    return "\n\n".join(
        f"--- Page {i + 1} ---\n{text}" for i, text in enumerate(page_texts) if text.strip()
    )


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


def _extract_single(pdf_path, api_key, vision_model, base_url, force_vision):
    """Extract text from one PDF; returns (filename, text)."""
    filename = os.path.basename(pdf_path)
    if force_vision and api_key:
        text = extract_pdf_with_vision(pdf_path, api_key, vision_model, base_url)
    elif _has_text(pdf_path):
        if api_key:
            text = extract_pdf_hybrid(pdf_path, api_key, vision_model, base_url)
        else:
            text = extract_pdf_text(pdf_path)
    elif api_key:
        text = extract_pdf_with_vision(pdf_path, api_key, vision_model, base_url)
    else:
        text = "[OCR ERROR] Scanned PDF detected but no API key for vision OCR"
    return filename, text


def extract_all_pdfs(
    folder_path: str,
    api_key: str = "",
    vision_model: str = "",
    base_url: str = "",
    progress_callback=None,
    force_vision: bool = False,
    max_workers: int = 3,
    **kwargs,
) -> dict[str, str]:
    """
    Extract text from all PDFs in a folder, processing them in parallel.

    When force_vision=True, always uses the vision LLM for all PDFs.
    max_workers controls how many PDFs are processed concurrently.

    Returns a dict mapping filename to extracted text.
    """
    pdf_files = sorted(glob.glob(os.path.join(folder_path, "*.pdf")))
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in {folder_path}")

    total = len(pdf_files)
    results = {}
    completed = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_path = {
            executor.submit(
                _extract_single, pdf_path, api_key, vision_model, base_url, force_vision
            ): pdf_path
            for pdf_path in pdf_files
        }
        for future in as_completed(future_to_path):
            try:
                filename, text = future.result()
            except Exception as e:
                filename = os.path.basename(future_to_path[future])
                text = f"[OCR ERROR] {str(e)}"
            results[filename] = text
            completed += 1
            if progress_callback:
                progress_callback(completed, total, filename)

    # Return in original sorted order
    return {os.path.basename(p): results[os.path.basename(p)] for p in pdf_files}
