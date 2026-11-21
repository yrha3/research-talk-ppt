---
name: research-talk-ppt
description: Create editable Chinese or English PowerPoint decks for scientific talks, lab meetings, journal clubs, and research progress updates. Use for 科研汇报 PPT, 组会, 文献汇报, or a figure-led paper presentation based on PDFs, notes, figures, or experiment records.
---

# Research Talk PPT

Turn the scientific argument into an editable 16:9 deck. Default to the user's language, audience and timing. Follow an explicitly requested visual template when supplied.

1. Read the evidence. Prefer the paper and figures for claims and the latest verified run artifacts for the user's own results. Check central values, cohorts, baselines and metric directions. DeepPaperNote is a useful index.
2. Read `references/layouts.md`. For each main slide, choose the question, the evidence the audience should look at first, and the answer they can take away. Plan the sequence from question to mechanism, comparison and boundary. Use an instructive failure or difficult example when available and give it a separate page when it answers a different question from the aggregate result.
3. Build the page around the chosen figure or diagram. Crop to the relevant panels, retaining labels, axes, uncertainty and legends. If details are too small for projection, split the figures along the argument. Let the slide title or a short sentence under the evidence state the conclusion once. Put narration in notes and keep qualifiers visible when they change the meaning of the claim.
4. Create a UTF-8 JSON manifest following `examples/demo.json`. Use a white canvas, dark main text, quiet sources and one optional highlight per slide. Preserve paper colors that encode scientific information. Use the local renderer for a local PPTX with Python 3.10+ and `python-pptx>=1.0.2,<2`. Resolve the script relative to this skill directory.
   `python scripts/build_deck.py path/to/deck.json path/to/output.pptx`
5. Render with PowerPoint or LibreOffice. Scan the deck as thumbnails to check that each slide has a clear focal point, then inspect every slide at presentation size for readable axes, labels, glyphs and footers. Reopen the PPTX to verify notes and separate image objects. Report any unavailable verification.

Use only materials authorized for the requested audience. For public examples use original or appropriately licensed assets with attribution. Keep source figures as separate image objects. Use native editable text and shapes for the narrative and simple diagrams.
