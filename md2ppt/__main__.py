"""
md2ppt.__main__
---------------
CLI entry point.  Invoked via:

    python -m md2ppt input.md output.pptx
    md2ppt input.md output.pptx          # after pip install
"""

import sys
from .slides_builder import SlidesBuilder

# Default ETSU color theme — passed through when the CLI is used directly.
# Users calling SlidesBuilder programmatically can supply their own palette.
_ETSU_COLORS = {
    "dk1":     "000000",   # Body text — black
    "lt1":     "FFFFFF",   # Slide background — white
    "dk2":     "00053E",   # Deepest navy
    "lt2":     "FFC72C",   # Bright gold — decorative/line use only
    "accent1": "003865",   # ETSU Blue          contrast ~8.5:1
    "accent2": "005A9E",   # Medium blue        contrast ~6.1:1
    "accent3": "7A6000",   # Dark amber/bronze  contrast ~7.2:1
    "accent4": "00053E",   # Deep navy repeat   contrast ~16:1
    "accent5": "1A5276",   # Steel blue         contrast ~7.8:1
    "accent6": "4A4A4A",   # Charcoal           contrast ~9.7:1
    "hlink":   "003865",   # Hyperlinks in ETSU blue
}


def main() -> None:
    """Console-script entry point registered in pyproject.toml."""
    if len(sys.argv) != 3:
        print("Usage: md2ppt <input.md> <output.pptx>")
        sys.exit(1)

    md_path, out_path = sys.argv[1], sys.argv[2]

    try:
        with open(md_path, "r", encoding="utf-8") as fh:
            markdown_text = fh.read()
    except FileNotFoundError:
        print(f"Error: file not found — {md_path}")
        sys.exit(1)

    SlidesBuilder(markdown_text, out_path, theme_colors=_ETSU_COLORS).build()


if __name__ == "__main__":
    main()
