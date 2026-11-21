---
name: research-talk-ppt
description: Create editable Chinese or English PowerPoint decks for scientific talks, lab meetings, journal clubs, and research progress updates. Use for 科研汇报 PPT, 组会, 文献汇报, or a figure-led paper presentation based on PDFs, notes, figures, or experiment records.
---

# Research Talk PPT

Turn the scientific argument into an editable 16:9 deck. Default to the user's language and stated duration. If neither duration nor slide count is given, plan about 10–13 slides for 15 minutes. Use the local renderer when the requested deliverable is a local PPTX.

1. Read the supplied evidence. Prefer the paper and figures for claims and the latest verified run artifacts for the user's own results. DeepPaperNote output is a useful index. Check central values, cohorts, baselines and metric directions against their sources.
2. Draft a slide sequence around the scientific question, source of evidence, mechanism, strongest comparison, failure case, and limitations. Adapt this order for progress talks or short talks. Use `references/layouts.md` for page selection and evidence handling.
3. Prepare UTF-8 JSON following `examples/demo.json`. For figure pages, supply visually checked crops and preserve panel letters, axes, uncertainty and legends. Resolve image paths from the manifest directory. Place one supported takeaway on an evidence page, with a source locator and speaker notes.
4. Build using Python 3.10+ with `python-pptx>=1.0.2,<2`. Resolve the script relative to this skill directory.
   `python scripts/build_deck.py path/to/deck.json path/to/output.pptx`
5. Reopen the PPTX to verify slides, notes and media. Render with available PowerPoint or LibreOffice and inspect every slide at presentation size. Check Chinese glyphs, long titles, cropped panels, evidence footers and readability. Report any verification unavailable. Return the deck and optionally the editable manifest when requested.

Use only materials authorized for the requested audience. For public examples use original or appropriately licensed assets with attribution. Keep source figures as separate image objects. Use native editable text and shapes for the narrative and simple diagrams.
