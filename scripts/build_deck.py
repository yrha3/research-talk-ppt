"""Render a research-talk JSON manifest to editable PowerPoint."""
import argparse
import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


INK = "162A36"
ACCENT = "006F69"
MUTED = "415765"
PAPER = "FAFBFA"
TINT = "E8F3F0"
FONT = "Microsoft YaHei"


def color(hex_value):
    return RGBColor.from_string(hex_value)


def box(slide, x, y, width, height, fill):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color(fill)
    shape.line.fill.background()
    return shape


def text(slide, value, x, y, width, height, size, ink=INK, bold=False):
    frame = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(width), Inches(height)
    ).text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Pt(0)
    frame.margin_top = frame.margin_bottom = Pt(0)
    frame.text = value
    for paragraph in frame.paragraphs:
        paragraph.font.name = FONT
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color(ink)


def picture(slide, path, x, y, width, height):
    shape = slide.shapes.add_picture(str(path), Inches(x), Inches(y))
    scale = min(Inches(width) / shape.width, Inches(height) / shape.height)
    shape.width = round(shape.width * scale)
    shape.height = round(shape.height * scale)
    shape.left = Inches(x) + (Inches(width) - shape.width) // 2
    shape.top = Inches(y) + (Inches(height) - shape.height) // 2


def draw_figure(slide, item, directory):
    images = item["images"]
    if not 1 <= len(images) <= 2:
        raise ValueError("A figure slide needs one or two images")
    width = 12.17 if len(images) == 1 else 5.9
    for index, image in enumerate(images):
        x = 0.58 + index * 6.27
        picture(slide, directory / image["path"], x, 1.32, width, 4.45)
        text(slide, image.get("caption", ""), x, 5.82, width, 0.34, 13, MUTED)


def draw_flow(slide, steps):
    if not 2 <= len(steps) <= 5:
        raise ValueError("A flow slide needs two to five steps")
    width = (12.17 - (len(steps) - 1) * 0.28) / len(steps)
    for index, step in enumerate(steps):
        x = 0.58 + index * (width + 0.28)
        box(slide, x, 2.1, width, 2.7, TINT)
        text(slide, step["label"], x + 0.18, 2.48, width - 0.36, 0.91, 22, ACCENT, True)
        text(slide, step["body"], x + 0.18, 3.48, width - 0.36, 1.12, 16)
        if index < len(steps) - 1:
            text(slide, "→", x + width + 0.02, 3.12, 0.28, 0.42, 15, ACCENT)


def draw_points(slide, points):
    if not 2 <= len(points) <= 4:
        raise ValueError("A points slide needs two to four points")
    for index, point in enumerate(points):
        y = 1.58 + index * 1.13
        box(slide, 0.58, y + 0.03, 0.06, 0.74, ACCENT)
        text(slide, point, 0.84, y, 11.9, 0.86, 20)


def draw_slide(deck, item, directory, number):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color(PAPER)
    kind = item["type"]
    if kind in {"cover", "closing"}:
        box(slide, 0.72, 1.75, 0.12, 3.38, ACCENT)
        text(slide, item["title"], 1.02, 2.05, 11.1, 2.0, 38, INK, True)
        text(slide, item.get("subtitle", ""), 1.04, 4.34, 10.9, 0.84, 19, MUTED)
    else:
        text(slide, item["title"], 0.58, 0.36, 12.17, 0.82, 27, INK, True)
        if kind == "figure":
            draw_figure(slide, item, directory)
        elif kind == "flow":
            draw_flow(slide, item["steps"])
        elif kind == "points":
            draw_points(slide, item["points"])
        else:
            raise ValueError(f"Unknown slide type: {kind}")
        if item.get("takeaway"):
            box(slide, 0.58, 6.24, 12.17, 0.52, TINT)
            text(slide, item["takeaway"], 0.74, 6.34, 11.83, 0.31, 15, ACCENT, True)
    text(slide, item.get("source", ""), 0.58, 7.05, 11.4, 0.28, 10, MUTED)
    text(slide, str(number), 12.26, 7.05, 0.5, 0.28, 10, MUTED)
    if item.get("notes"):
        slide.notes_slide.notes_text_frame.text = item["notes"]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.output.suffix.lower() != ".pptx":
        parser.error("Output must have a .pptx extension")
    data = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    slides = data["slides"]
    if not isinstance(slides, list) or not slides:
        parser.error("The manifest must contain a nonempty slides list")
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.3333), Inches(7.5)
    for number, item in enumerate(slides, start=1):
        draw_slide(deck, item, args.manifest.parent, number)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    deck.save(str(args.output))


if __name__ == "__main__":
    main()
