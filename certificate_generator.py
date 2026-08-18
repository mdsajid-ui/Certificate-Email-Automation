"""
certificate_generator.py
-------------------------
Handles rendering a participant's name onto a certificate template
(PNG or PDF) and saving the result as an individual PDF certificate.
"""

import os
import re
from PIL import Image, ImageDraw, ImageFont

try:
    import fitz  # PyMuPDF - only needed if template is a PDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


def sanitize_filename(name: str) -> str:
    """Turn a person's name into a safe filename fragment."""
    name = str(name).strip()
    name = re.sub(r"[^\w\s-]", "", name)   # drop punctuation
    name = re.sub(r"\s+", "_", name)       # spaces -> underscores
    return name or "Participant"


def load_template_as_image(template_path: str) -> Image.Image:
    """Load a PNG/JPG template, or rasterize page 1 of a PDF template."""
    ext = os.path.splitext(template_path)[1].lower()
    if ext == ".pdf":
        if not HAS_FITZ:
            raise RuntimeError(
                "PyMuPDF (pymupdf) is required to use a PDF certificate template. "
                "Install it with: pip install pymupdf"
            )
        doc = fitz.open(template_path)
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x for better resolution
        img_bytes = pix.tobytes("png")
        doc.close()
        from io import BytesIO
        return Image.open(BytesIO(img_bytes)).convert("RGB")
    else:
        return Image.open(template_path).convert("RGB")


def get_font(font_path: str, size: int) -> ImageFont.FreeTypeFont:
    """Load a TTF font, falling back to PIL's default bitmap font if unavailable."""
    try:
        if font_path and os.path.exists(font_path):
            return ImageFont.truetype(font_path, size)
        # Try a couple of common system fonts as a fallback
        for candidate in [
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ]:
            if os.path.exists(candidate):
                return ImageFont.truetype(candidate, size)
    except Exception:
        pass
    return ImageFont.load_default()


def render_certificate_image(
    template_img: Image.Image,
    name: str,
    font_path: str = None,
    font_size: int = 60,
    text_color: tuple = (5, 3, 116),
    y_position_pct: float = 0.50,
) -> Image.Image:
    """
    Return a copy of the template image with `name` drawn centered
    horizontally at `y_position_pct` (0.0 = top, 1.0 = bottom) of the image.
    """
    img = template_img.copy()
    draw = ImageDraw.Draw(img)
    font = get_font(font_path, font_size)

    text = str(name).strip()
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    x = (img.width - text_w) / 2
    y = (img.height * y_position_pct) - (text_h / 2)

    draw.text((x, y), text, font=font, fill=text_color)
    return img


def save_image_as_pdf(img: Image.Image, output_path: str):
    """Save a PIL image as a single-page PDF."""
    rgb_img = img.convert("RGB")
    rgb_img.save(output_path, "PDF", resolution=150.0)


def generate_certificate(
    template_img: Image.Image,
    name: str,
    output_dir: str,
    font_path: str = None,
    font_size: int = 60,
    text_color: tuple = (5, 3, 116),
    y_position_pct: float = 0.50,
    used_names: dict = None,
) -> str:
    """
    Generate one certificate for `name` and save it to `output_dir`.
    Returns the output file path. Handles duplicate names by suffixing _2, _3, etc.
    """
    os.makedirs(output_dir, exist_ok=True)
    used_names = used_names if used_names is not None else {}

    rendered = render_certificate_image(
        template_img, name, font_path, font_size, text_color, y_position_pct
    )

    base = sanitize_filename(name)
    used_names[base] = used_names.get(base, 0) + 1
    suffix = "" if used_names[base] == 1 else f"_{used_names[base]}"
    filename = f"{base}{suffix}_Certificate.pdf"
    output_path = os.path.join(output_dir, filename)

    save_image_as_pdf(rendered, output_path)
    return output_path
