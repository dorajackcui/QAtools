"""Export the generated Q artwork as rounded, transparent Windows icon assets."""
from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)


def main() -> None:
    with Image.open(ROOT / "packaging/QAtools-icon.png") as source:
        artwork = source.convert("RGBA")
    width, height = artwork.size
    if width != height:
        raise ValueError("The icon artwork must be square")
    # Supersampling keeps curved edges smooth at every Windows icon size.
    scale = 4
    with Image.new("L", (width * scale, height * scale), 0) as large_mask:
        ImageDraw.Draw(large_mask).rounded_rectangle(
            (0, 0, width * scale - 1, height * scale - 1),
            radius=width * scale * 0.18,
            fill=255,
        )
        with large_mask.resize(artwork.size, Image.Resampling.LANCZOS) as mask:
            artwork.putalpha(mask)
    try:
        artwork.save(ROOT / "packaging/QAtools-icon-rounded.png")
        artwork.save(
            ROOT / "packaging/QAtools.ico", format="ICO",
            sizes=[(size, size) for size in SIZES],
        )
    finally:
        artwork.close()


if __name__ == "__main__":
    main()
