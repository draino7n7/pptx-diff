# pptx-diff

A pure-Python command-line and file-picker utility that extracts text and image representations of sequential PowerPoint (`.pptx`) training versions, compares them, and generates a compact 3–4 line summary per version transition designed to be copy-pasted directly into a single Excel spreadsheet cell.

## Features

- **Interactive Multi-Version File Picker (or CLI Arguments):**
  - Run `python pptx_diff.py` with no arguments to open a native Windows file picker for Version 1, followed by Version 2 (starting automatically in the same folder), and a prompt asking whether you want to add Version 3, Version 4, etc.
  - Or pass 2+ `.pptx` paths directly on the command line.
- **Structured Extraction (`output/`):**
  - Creates a dedicated timestamped run directory inside `output/` (ignored by git).
  - Extracts all slide images into `<version>_images/` and writes a structured `<version>_extracted.txt` linking slide titles, categories, body text, tables, charts, clipart, on-slide references, speaker notes, and extracted images.
- **Training Deck Structure Awareness:**
  - Automatically identifies **Title Slide**, **Course Terminal Objectives**, **Module Terminal Objectives**, **Lesson Learning Objective**, **Section Enabling Objective**, **References**, **Knowledge Check**, and **Content** slides.
  - Parses the main **References** slide first and uses its entries (combined with bottom-left slide position heuristics) to recognize and track on-slide reference text boxes separately from body text.
  - Tracks **Speaker Notes** updates across all slides.
- **Visual & Formatting Intelligence:**
  - Compares images using both SHA-256 and perceptual difference hashing (`dHash`) so moving or slightly resizing an image on a slide does **not** trigger a false image change (rolling small layout/font tweaks into a single `Minor formatting changes` line).
  - Aligns slides across versions by internal slide ID and content similarity so added or deleted slides do not shift and falsely flag subsequent slides.

## Setup

```powershell
pip install -r requirements.txt
```

## Usage

### 1. Interactive File Picker Mode (Recommended)
Run from PowerShell:
```powershell
python pptx_diff.py
```
1. Pick the first `.pptx` version in the file dialog.
2. Pick the next `.pptx` version in the second file dialog.
3. Click **Yes** to keep adding versions (`v2`, `v3`, ...) for the same training, or **No** to finish and run the comparison.
4. When finished, `summary.txt` automatically opens in Notepad ready to copy-paste each version block into Excel.

### 2. Command-Line Arguments Mode
```powershell
python pptx_diff.py path\to\Section_v0.pptx path\to\Section_v1.pptx path\to\Section_v2.pptx
```

Optional flags:
- `--output-dir <path>`: Specify a custom output directory instead of `./output`.
- `--no-open`: Do not automatically open `summary.txt` when finished.

## Output Files

Each run creates a folder `output/<deck_name>_<YYYYMMDD_HHMMSS>/` containing:
- `summary.txt`: High-level 3–4 line summary per version transition (Option B format) ready to copy-paste into an Excel cell.
- `detailed_diff.txt`: Slide-by-slide breakdown of every detected change.
- `v1_<name>_extracted.txt`, `v2_<name>_extracted.txt`, ...: Full text representation of each presentation version.
- `v1_<name>_images/`, `v2_<name>_images/`, ...: Extracted images from each presentation version.

### Example `summary.txt` Output
```text
=== Comparison 1: SectionA_v0.pptx -> SectionA_v1.pptx ===
Slides: 1 added (8), 1 deleted (old 9) | Images changed: Slides 7, 8
Objectives/Refs: Section Enabling Objective (Slide 5) & References (Slides 6, 7, 8) updated
Knowledge Checks: Updated on Slide 9
Content/Notes: Speaker notes changed (7, 9)
```

## Running Tests

```powershell
python test_pptx_diff.py
```
