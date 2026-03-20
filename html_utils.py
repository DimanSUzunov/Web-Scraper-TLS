"""
UPDATED
Utility functions for handling HTML-specific tasks, such as inserting images into a Word document.
"""
import requests
import io
from docx.shared import Inches
from docx.image.exceptions import UnrecognizedImageError
from typing import Any
from config import DEFAULT_IMAGE_WIDTH

def _svg_to_png_stream(svg_data: bytes) -> io.BytesIO:
    """Converts SVG bytes to PNG using svglib + reportlab (Windows-friendly)."""
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPM
    import tempfile, os

    with tempfile.NamedTemporaryFile(suffix=".svg", delete=False) as f:
        f.write(svg_data)
        tmp_path = f.name
    try:
        drawing = svg2rlg(tmp_path)
        png_stream = io.BytesIO()
        renderPM.drawToFile(drawing, png_stream, fmt="PNG")
        png_stream.seek(0)
        return png_stream
    finally:
        os.unlink(tmp_path)

def _try_add_picture(doc, image_stream: io.BytesIO, src: str) -> None:
    """Attempts to insert an image, converting SVG to PNG if needed."""
    try:
        doc.add_paragraph().add_run().add_picture(image_stream, width=Inches(DEFAULT_IMAGE_WIDTH))
    except UnrecognizedImageError:
        # Try converting as SVG
        image_stream.seek(0)
        raw = image_stream.read()
        if raw.strip().startswith(b'<') or b'<svg' in raw[:512]:
            try:
                png_stream = _svg_to_png_stream(raw)
                doc.add_paragraph().add_run().add_picture(png_stream, width=Inches(DEFAULT_IMAGE_WIDTH))
                return
            except Exception:
                pass
        # Fallback: unsupported format
        p = doc.add_paragraph()
        run = p.add_run(f"[Afbeeldingsformaat niet ondersteund] - bron: {src}")
        run.italic = True

def insert_image(doc: Any, src: str) -> None:
    """
    Downloads an image from the given source URL and inserts it into the Word document.
    Supports JPEG, PNG, GIF, TIFF, BMP, and SVG (via cairosvg).
    Adds a caption with the image source. Handles errors gracefully.
    """
    try:
        img_response = requests.get(src, timeout=5)
        img_response.raise_for_status()
        image_stream = io.BytesIO(img_response.content)
        _try_add_picture(doc, image_stream, src)
        link = doc.add_paragraph().add_run(f"Afbeeldingsbron: {src}")
        link.italic = True
    except requests.exceptions.RequestException as e:
        p = doc.add_paragraph()
        run = p.add_run(f"[Afbeelding kon niet geladen worden] - bron: {src} ({e})")
        run.bold = True

def insert_inline_svg(doc: Any, svg_string: str) -> None:
    """
    Converts an inline SVG string to PNG and inserts it into the Word document.
    Falls back to a placeholder if conversion fails.
    """
    try:
        png_stream = _svg_to_png_stream(svg_string.encode("utf-8"))
        doc.add_paragraph().add_run().add_picture(png_stream, width=Inches(DEFAULT_IMAGE_WIDTH))
    except Exception as e:
        p = doc.add_paragraph()
        run = p.add_run(f"[Inline SVG kon niet worden weergegeven: {e}]")
        run.italic = True