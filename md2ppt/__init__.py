"""
md2ppt
------
Convert Markdown text into a .pptx PowerPoint presentation.

Quick start::

    from md2ppt import SlidesBuilder

    with open("slides.md") as f:
        md = f.read()

    SlidesBuilder(md, "output.pptx").build()
"""

from .slides_builder import SlidesBuilder, SlideData, SlideItem

__all__ = ["SlidesBuilder", "SlideData", "SlideItem"]
__version__ = "0.1.0"
