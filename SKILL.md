---
name: research-talk-ppt
description: Create editable Chinese or English PowerPoint decks for scientific talks, lab meetings, journal clubs, and research progress updates. Use for 科研汇报 PPT, 组会, 文献汇报, or a figure-led paper presentation based on PDFs, notes, figures, or experiment records.
---

# Research Talk PPT

Turn the scientific argument into an editable 16:9 deck. Default to a white canvas, centered dark titles, one short burnt-orange title emphasis or figure callout, and large, centered evidence. Follow the user's language, audience, timing and any supplied visual template.

1. Read the evidence. Prefer the paper and figures for claims and the latest verified run artifacts for the user's own results. Check central values, cohorts, baselines and metric directions. DeepPaperNote is a useful index.
2. Read `references/layouts.md`. For each main slide, choose the question, the evidence the audience should look at first, and the answer they can take away. In an integrated talk, merge duplicated introductions and use a concrete unresolved question to connect each section to the next. Translate a source talk's questions when requested. Give a difficult example its own page when it answers a different question from the aggregate result.
3. Build the page around one figure, aligned comparison or editable table. Crop to the relevant panels, retaining labels, axes, uncertainty and legends. Split dense evidence across pages at a readable size. Use a title or a short sentence near the evidence to state the conclusion once, keeping qualifications that change its meaning visible. Put the reading order, spoken transition and full citations in notes. When timing matters, budget the notes across the whole talk.
   Give every slide that uses a published claim, figure, table, or dataset a short citation in the bottom-left footer using `source`. Use one line of small gray text, typically author and year plus the relevant figure, table, or section. Put full reference details in `notes` or in a references slide when requested. Keep the footer free of colored backgrounds and decoration.
4. Create a UTF-8 JSON manifest following `examples/demo.json`. Use a white canvas, dark main text, quiet sources and one optional highlight per slide. Preserve paper colors that encode scientific information. Use the local renderer for a local PPTX with Python 3.10+ and `python-pptx>=1.0.2,<2`. Resolve the script relative to this skill directory.
   `python scripts/build_deck.py path/to/deck.json path/to/output.pptx`
5. Render with PowerPoint or LibreOffice. Scan the deck as thumbnails to check that each slide has a clear focal point, then inspect every slide at presentation size for readable axes, labels, glyphs and footers. Check that citations fit on one line and identify the evidence on that slide. Reopen the PPTX to verify notes and separate image objects. Report any unavailable verification.

Use only materials authorized for the requested audience. For public examples use original or appropriately licensed assets with attribution. Keep source figures as separate image objects. Use native editable text and shapes for the narrative and simple diagrams.
