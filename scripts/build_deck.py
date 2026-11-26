"""Render a research-talk JSON manifest to editable PowerPoint."""
import argparse
import json
from math import isfinite
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


INK = "172026"
ACCENT = "A94F13"
COVER_INK = "193A56"
MUTED = "505960"
PAPER = "FFFFFF"
TABLE_HEAD = "EAF0F4"
RULE = "D8DEE3"
FONT = "Microsoft YaHei"
MARGIN = 0.58
CONTENT_W = 12.17


def color(hex_value):
    return RGBColor.from_string(hex_value)


def box(slide, x, y, width, height, fill):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(width), Inches(height)
    )
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = color(fill)
    else:
        shape.fill.background()
    return shape


def east_asian_font(paragraph):
    east_asian = OxmlElement("a:ea")
    east_asian.set("typeface", FONT)
    paragraph._p.get_or_add_pPr().get_or_add_defRPr().append(east_asian)


def text(slide, value, x, y, width, height, size, ink=INK, bold=False,
         align=PP_ALIGN.LEFT, emphasis=""):
    if not value:
        return
    frame = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(width), Inches(height)
    ).text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Pt(0)
    frame.margin_top = frame.margin_bottom = Pt(0)
    frame.text = value
    for paragraph in frame.paragraphs:
        paragraph.alignment = align
        paragraph.space_before = paragraph.space_after = Pt(0)
        paragraph.line_spacing = 1.08
        paragraph.font.name = FONT
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color(ink)
        # Explicit East Asian fallback keeps CJK text independent of the slide theme.
        east_asian_font(paragraph)
        if emphasis and emphasis in paragraph.text:
            before, _, after = paragraph.text.partition(emphasis)
            paragraph.clear()
            paragraph.add_run().text = before
            focus = paragraph.add_run()
            focus.text = emphasis
            focus.font.color.rgb = color(ACCENT)
            paragraph.add_run().text = after
            emphasis = ""


def picture(slide, path, x, y, width, height):
    shape = slide.shapes.add_picture(str(path), Inches(x), Inches(y))
    scale = min(Inches(width) / shape.width, Inches(height) / shape.height)
    shape.width = round(shape.width * scale)
    shape.height = round(shape.height * scale)
    shape.left = Inches(x) + (Inches(width) - shape.width) // 2
    shape.top = Inches(y) + (Inches(height) - shape.height) // 2
    return shape


def outline_focus(slide, image, bounds):
    x, y, width, height = bounds
    if not (0 <= x < 1 and 0 <= y < 1 and width > 0 and height > 0
            and x + width <= 1 and y + height <= 1):
        raise ValueError("Highlight must lie within the image in normalized coordinates")
    marker = box(slide, (image.left + image.width * x) / Inches(1),
                 (image.top + image.height * y) / Inches(1),
                 image.width * width / Inches(1), image.height * height / Inches(1), None)
    marker.line.color.rgb = color(ACCENT)
    marker.line.width = Pt(1.5)


def draw_figure(slide, item, directory):
    images = item["images"]
    if not 1 <= len(images) <= 2:
        raise ValueError("A figure slide needs one or two images")
    layout = item.get("layout", "row")
    if layout not in {"row", "stack"}:
        raise ValueError("Figure layout must be row or stack")
    stacked = layout == "stack" and len(images) == 2
    top = 1.42
    bottom = 6.16 if item.get("takeaway") else 6.86
    width = CONTENT_W if stacked or len(images) == 1 else (CONTENT_W - 0.35) / 2
    block_height = (bottom - top - 0.18) / 2 if stacked else bottom - top
    for index, image in enumerate(images):
        x = MARGIN if stacked else MARGIN + index * (width + 0.35)
        y = top + index * (block_height + 0.18) if stacked else top
        caption = image.get("caption", "")
        image_height = block_height - (0.31 if caption else 0)
        shape = picture(slide, directory / image["path"], x, y, width, image_height)
        if "highlight" in image:
            outline_focus(slide, shape, image["highlight"])
        caption_y = (shape.top + shape.height) / Inches(1)
        text(slide, caption, x, caption_y + 0.05,
             width, 0.26, 12, MUTED, align=PP_ALIGN.CENTER)


def draw_flow(slide, steps):
    if not 2 <= len(steps) <= 5:
        raise ValueError("A flow slide needs two to five steps")
    gap = 0.52
    width = (CONTENT_W - (len(steps) - 1) * gap) / len(steps)
    for index, step in enumerate(steps):
        x = MARGIN + index * (width + gap)
        text(slide, step["label"], x, 2.68, width, 0.8, 23, INK, True, PP_ALIGN.CENTER)
        text(slide, step["body"], x, 3.63, width, 1.5, 17, MUTED, align=PP_ALIGN.CENTER)
        if index < len(steps) - 1:
            text(slide, "→", x + width, 2.71, gap, 0.42, 18, MUTED, align=PP_ALIGN.CENTER)


def draw_points(slide, points):
    if not 2 <= len(points) <= 4:
        raise ValueError("A points slide needs two to four points")
    if all(isinstance(point, dict) for point in points):
        for index, point in enumerate(points):
            y = 1.66 + index * 1.10
            if index:
                box(slide, 1.22, y - 0.15, 10.85, 0.01, RULE)
            text(slide, point["label"], 1.22, y, 2.48, 0.72, 18, ACCENT, True)
            text(slide, point["body"], 3.97, y, 8.10, 0.94, 19)
        return
    if not all(isinstance(point, str) for point in points):
        raise ValueError("Points must all be strings or all be label/body objects")
    for index, point in enumerate(points):
        y = 1.90 + index * 1.00
        text(slide, "•", 1.22, y, 0.28, 0.48, 19, MUTED)
        text(slide, point, 1.65, y, 10.45, 0.85, 22)


def draw_table(slide, item):
    rows = item["rows"]
    if not isinstance(rows, list) or len(rows) < 2 or not isinstance(rows[0], list) or not rows[0]:
        raise ValueError("A table needs a nonempty header row and at least one data row")
    columns = len(rows[0])
    if any(not isinstance(row, list) or len(row) != columns for row in rows):
        raise ValueError("All table rows must have the same number of cells")
    widths = item.get("column_widths", [1] * columns)
    if not isinstance(widths, list) or len(widths) != columns or any(
        not isinstance(width, (int, float)) or not isfinite(width) or width <= 0 for width in widths
    ):
        raise ValueError("column_widths needs one positive weight per column")

    lead = item.get("lead", "")
    top = 1.81 if lead else 1.45
    bottom = 6.13 if item.get("takeaway") else 6.53
    height = min(bottom - top, 0.77 * len(rows))
    y = top + (bottom - top - height) / 2
    text(slide, lead, MARGIN, y - 0.53, CONTENT_W, 0.45, 15, MUTED, align=PP_ALIGN.CENTER)
    table = slide.shapes.add_table(
        len(rows), columns, Inches(MARGIN), Inches(y), Inches(CONTENT_W), Inches(height)
    ).table
    table.first_row = table.horz_banding = False
    total = sum(widths)
    for column, width in zip(table.columns, widths):
        column.width = Inches(CONTENT_W * (width / total))

    for index, values in enumerate(rows):
        row = table.rows[index]
        row.height = Inches(height / len(rows))
        for cell, value in zip(row.cells, values):
            if not isinstance(value, (str, int, float)):
                raise ValueError("Table cells must contain text or a number")
            cell.text = str(value)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Pt(1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = color(TABLE_HEAD if index == 0 else PAPER)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.space_before = paragraph.space_after = Pt(0)
                paragraph.line_spacing = 1.0
                paragraph.font.name = FONT
                paragraph.font.size = Pt(14 if len(rows) >= 8 else 16)
                paragraph.font.bold = index == 0
                paragraph.font.color.rgb = color(INK)
                east_asian_font(paragraph)


def draw_slide(deck, item, directory, number):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color(PAPER)
    kind = item["type"]
    title = item["title"]
    if not isinstance(title, str) or not title.strip():
        raise ValueError("Slide title must be a nonempty string")
    emphasis = item.get("emphasis", "")
    highlights = sum("highlight" in image for image in item["images"]) if kind == "figure" else 0
    if highlights + bool(emphasis) > 1:
        raise ValueError("Use at most one highlight or title emphasis per slide")
    if emphasis and not any(emphasis in line for line in title.splitlines()):
        raise ValueError("Title emphasis must be a literal phrase on one line of the title")
    if kind in {"cover", "closing"}:
        text(slide, title, 1.00, 2.17, 11.33, 1.84, 39, COVER_INK if kind == "cover" else INK, True,
             PP_ALIGN.CENTER, emphasis)
        text(slide, item.get("subtitle", ""), 1.00, 4.22, 11.33, 0.85,
             18, MUTED, align=PP_ALIGN.CENTER)
    else:
        title_size = 24 if "\n" in title else 30
        text(slide, title, MARGIN, 0.41, CONTENT_W, 0.92, title_size, INK, True,
             PP_ALIGN.CENTER, emphasis)
        if kind == "figure":
            draw_figure(slide, item, directory)
        elif kind == "table":
            draw_table(slide, item)
        elif kind == "flow":
            draw_flow(slide, item["steps"])
        elif kind == "points":
            draw_points(slide, item["points"])
        else:
            raise ValueError(f"Unknown slide type: {kind}")
        text(slide, item.get("takeaway", ""), MARGIN, 6.37, CONTENT_W, 0.54,
             16, INK, align=PP_ALIGN.CENTER)
    text(slide, item.get("source", ""), MARGIN, 7.10, 11.4, 0.25, 10, MUTED)
    if kind not in {"cover", "closing"}:
        text(slide, str(number), 12.25, 7.10, 0.5, 0.25, 10, MUTED, align=PP_ALIGN.RIGHT)
    if item.get("notes"):
        slide.notes_slide.notes_text_frame.text = item["notes"]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.output.suffix.lower() != ".pptx":
        parser.error("Output must have a .pptx extension")
    try:
        data = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        parser.error(f"{args.manifest}: {exc}")
    slides = data.get("slides") if isinstance(data, dict) else None
    if not isinstance(slides, list) or not slides:
        parser.error("The manifest must contain a nonempty slides list")
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.3333), Inches(7.5)
    for number, item in enumerate(slides, start=1):
        try:
            draw_slide(deck, item, args.manifest.parent, number)
        except KeyError as exc:
            parser.error(f"Slide {number}: missing field {exc}")
        except (AttributeError, OSError, TypeError, ValueError) as exc:
            parser.error(f"Slide {number}: {exc}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    deck.save(str(args.output))


if __name__ == "__main__":
    main()
