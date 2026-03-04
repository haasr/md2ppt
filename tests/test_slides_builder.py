"""
Basic smoke tests for md2ppt.SlidesBuilder.
Run with: pytest tests/
"""

import os
import tempfile
import pytest
from md2ppt import SlidesBuilder
from md2ppt.slides_builder import SlideData, SlideItem


SAMPLE_MD = """\
# Test Presentation: A Subtitle

## Bullet Slide

- First bullet
- Second bullet

## Numbered Slide

1. Item one
2. Item two

## Plain Text Slide

This is plain body text.
No bullets here.
"""


def test_parse_produces_correct_slide_count():
    builder = SlidesBuilder(SAMPLE_MD, "dummy.pptx")
    slides = builder._parse_markdown()
    assert len(slides) == 4


def test_title_slide_parsed():
    builder = SlidesBuilder(SAMPLE_MD, "dummy.pptx")
    slides = builder._parse_markdown()
    title_slide = slides[0]
    assert title_slide.kind == SlideData.TITLE_SLIDE
    assert title_slide.title == "Test Presentation"
    assert title_slide.subtitle == "A Subtitle"


def test_bullet_items_parsed():
    builder = SlidesBuilder(SAMPLE_MD, "dummy.pptx")
    slides = builder._parse_markdown()
    bullet_slide = slides[1]
    assert len(bullet_slide.items) == 2
    assert all(item.kind == SlideItem.BULLET for item in bullet_slide.items)


def test_numbered_items_parsed():
    builder = SlidesBuilder(SAMPLE_MD, "dummy.pptx")
    slides = builder._parse_markdown()
    numbered_slide = slides[2]
    assert len(numbered_slide.items) == 2
    assert all(item.kind == SlideItem.NUMBERED for item in numbered_slide.items)
    assert numbered_slide.items[0].number == 1
    assert numbered_slide.items[1].number == 2


def test_plain_items_parsed():
    builder = SlidesBuilder(SAMPLE_MD, "dummy.pptx")
    slides = builder._parse_markdown()
    plain_slide = slides[3]
    assert len(plain_slide.items) == 2
    assert all(item.kind == SlideItem.PLAIN for item in plain_slide.items)


def test_build_creates_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        out = os.path.join(tmpdir, "test_output.pptx")
        SlidesBuilder(SAMPLE_MD, out).build()
        assert os.path.exists(out)
        assert os.path.getsize(out) > 0


def test_output_path_extension_appended():
    with tempfile.TemporaryDirectory() as tmpdir:
        out = os.path.join(tmpdir, "no_extension")
        SlidesBuilder(SAMPLE_MD, out).build()
        assert os.path.exists(out + ".pptx")
