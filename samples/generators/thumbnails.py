"""Crops and scales the exports thumbnails.ps1 made into the gallery's 800 x 450 WebP thumbnails.

    python thumbnails.py <output folder> <id>|<export.png or .pdf>|<anchor>|<page> ...

A PDF export is rasterised with PyMuPDF (the page asked for); the picture is then scaled to cover
16:9 and cropped at the anchor (top, center, bottom, or a fraction from 0 for the top to 1 for
the bottom), so a portrait page shows its head and a landscape drawing its middle. Needs
`pip install pymupdf pillow`.
"""

import io
import sys
from pathlib import Path

import fitz
from PIL import Image

WIDTH, HEIGHT = 800, 450


def load(path: Path, number: int) -> Image.Image:
    if path.suffix.lower() == ".pdf":
        with fitz.open(path) as document:
            page = document[number - 1]
            # About 1600 pixels across, whatever the page size.
            zoom = 1600 / page.rect.width
            pixels = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            return Image.open(io.BytesIO(pixels.tobytes("png"))).convert("RGB")
    picture = Image.open(path)
    if picture.mode in ("RGBA", "LA", "P"):
        # Exports with transparency sit on white, as they do on a page.
        picture = picture.convert("RGBA")
        white = Image.new("RGBA", picture.size, (255, 255, 255, 255))
        picture = Image.alpha_composite(white, picture)
    return picture.convert("RGB")


def cover(picture: Image.Image, anchor: str) -> Image.Image:
    scale = max(WIDTH / picture.width, HEIGHT / picture.height)
    size = (max(WIDTH, round(picture.width * scale)), max(HEIGHT, round(picture.height * scale)))
    scaled = picture.resize(size, Image.LANCZOS)
    left = (scaled.width - WIDTH) // 2
    spare = scaled.height - HEIGHT
    share = {"top": 0.0, "center": 0.5, "bottom": 1.0}.get(anchor)
    top = round(spare * (float(anchor) if share is None else share))
    return scaled.crop((left, top, left + WIDTH, top + HEIGHT))


def main() -> int:
    output = Path(sys.argv[1])
    for item in sys.argv[2:]:
        name, source, anchor, page = item.split("|")
        target = output / f"{name}.webp"
        cover(load(Path(source), int(page)), anchor).save(target, "WEBP", quality=82, method=6)
        print(f"{target.name}: {target.stat().st_size} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
