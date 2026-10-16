#!/usr/bin/env python3
"""
pptx_diff.py - PowerPoint (.pptx) Training Deck Version Comparison Tool

Extracts a structured text and image representation of two or more sequential
versions of a PowerPoint training presentation, compares them using pure Python
(no external AI calls), and produces a concise 3-4 line summary per version
transition ready to copy-paste into a single Excel spreadsheet cell.
"""

from __future__ import annotations

import argparse
import datetime
import difflib
import hashlib
import io
import os
import re
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

try:
    from PIL import Image
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False


# ============================================================================
# Constants & Slide Categories
# ============================================================================

CAT_TITLE = "Title Slide"
CAT_COURSE_TLO = "Course Terminal Objectives"
CAT_MODULE_TLO = "Module Terminal Objectives"
CAT_GENERIC_TLO = "Terminal Objectives"
CAT_LESSON_LO = "Lesson Learning Objective"
CAT_SECTION_EO = "Section Enabling Objective"
CAT_REFERENCES = "References"
CAT_KNOWLEDGE_CHECK = "Knowledge Check"
CAT_CONTENT = "Content"

OBJECTIVE_CATEGORIES = {
    CAT_COURSE_TLO,
    CAT_MODULE_TLO,
    CAT_GENERIC_TLO,
    CAT_LESSON_LO,
    CAT_SECTION_EO,
}

STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from", "are", "was", "were",
    "have", "has", "had", "not", "but", "what", "all", "when", "where", "which",
    "who", "will", "would", "can", "could", "should", "may", "might", "into",
    "onto", "upon", "about", "above", "below", "between", "through", "during",
    "before", "after", "under", "over", "again", "further", "then", "once",
    "here", "there", "why", "how", "any", "both", "each", "few", "more", "most",
    "other", "some", "such", "only", "own", "same", "than", "too", "very",
    "section", "lesson", "module", "course", "slide", "training",
}


# ============================================================================
# Data Structures
# ============================================================================

@dataclass
class TextBoxItem:
    shape_name: str
    text: str
    normalized_text: str
    norm_left: float
    norm_top: float
    norm_width: float
    norm_height: float
    format_signature: str
    is_title: bool = False
    is_reference: bool = False


@dataclass
class ImageItem:
    index_on_slide: int
    filename: str
    rel_path: str
    ext: str
    sha256: str
    dhash: Optional[int]
    byte_size: int
    norm_left: float
    norm_top: float
    norm_width: float
    norm_height: float
    format_signature: str


@dataclass
class TableItem:
    rows: List[List[str]]
    normalized_text: str
    format_signature: str


@dataclass
class ChartItem:
    chart_type: str
    title: str
    series_data: List[str]
    normalized_text: str
    format_signature: str


@dataclass
class ClipartItem:
    shape_type: str
    shape_name: str
    text: str
    format_signature: str


@dataclass
class SlideData:
    slide_number: int
    slide_id: int
    category: str
    title: str
    normalized_title: str
    text_boxes: List[TextBoxItem] = field(default_factory=list)
    body_text_items: List[str] = field(default_factory=list)
    on_slide_references: List[str] = field(default_factory=list)
    images: List[ImageItem] = field(default_factory=list)
    tables: List[TableItem] = field(default_factory=list)
    charts: List[ChartItem] = field(default_factory=list)
    cliparts: List[ClipartItem] = field(default_factory=list)
    diagram_texts: List[str] = field(default_factory=list)
    speaker_notes: str = ""
    normalized_notes: str = ""

    @property
    def combined_body_text(self) -> str:
        parts = list(self.body_text_items)
        if self.diagram_texts:
            parts.extend(self.diagram_texts)
        return "\n".join(parts).strip()

    @property
    def normalized_body_text(self) -> str:
        return normalize_text(self.combined_body_text)

    @property
    def normalized_refs_text(self) -> str:
        return normalize_text("\n".join(self.on_slide_references))

    @property
    def full_format_signature(self) -> str:
        sigs = (
            [tb.format_signature for tb in self.text_boxes]
            + [img.format_signature for img in self.images]
            + [tbl.format_signature for tbl in self.tables]
            + [cht.format_signature for cht in self.charts]
            + [clp.format_signature for clp in self.cliparts]
        )
        return "|".join(sigs)


@dataclass
class PresentationData:
    file_path: Path
    version_tag: str
    slide_width: int
    slide_height: int
    slides: List[SlideData] = field(default_factory=list)
    reference_catalog: List[str] = field(default_factory=list)
    extracted_text_path: Optional[Path] = None
    images_dir: Optional[Path] = None


@dataclass
class SlideDiffDetail:
    old_slide_num: Optional[int]
    new_slide_num: Optional[int]
    category: str
    title: str
    status: str  # "added", "deleted", "modified", "formatting_only", "unchanged"
    text_changed: bool = False
    notes_changed: bool = False
    refs_changed: bool = False
    images_added: int = 0
    images_deleted: int = 0
    images_replaced: int = 0
    tables_changed: bool = False
    charts_changed: bool = False
    clipart_changed: bool = False
    minor_formatting: bool = False
    change_notes: List[str] = field(default_factory=list)


@dataclass
class DeckComparisonResult:
    old_deck: PresentationData
    new_deck: PresentationData
    added_slides: List[SlideData] = field(default_factory=list)
    deleted_slides: List[SlideData] = field(default_factory=list)
    reordered_slides: List[Tuple[int, int]] = field(default_factory=list)
    image_changed_slides: List[int] = field(default_factory=list)
    images_added_slides: List[int] = field(default_factory=list)
    images_deleted_slides: List[int] = field(default_factory=list)
    images_replaced_slides: List[int] = field(default_factory=list)
    tables_changed_slides: List[int] = field(default_factory=list)
    charts_changed_slides: List[int] = field(default_factory=list)
    clipart_changed_slides: List[int] = field(default_factory=list)
    objectives_changed: List[Tuple[str, int]] = field(default_factory=list)
    main_ref_slides_changed: List[int] = field(default_factory=list)
    on_slide_refs_changed: List[int] = field(default_factory=list)
    kc_updated_slides: List[int] = field(default_factory=list)
    kc_added_slides: List[int] = field(default_factory=list)
    kc_deleted_old_slides: List[int] = field(default_factory=list)
    content_text_changed_slides: List[int] = field(default_factory=list)
    speaker_notes_changed_slides: List[int] = field(default_factory=list)
    minor_formatting_slides: List[int] = field(default_factory=list)
    slide_details: List[SlideDiffDetail] = field(default_factory=list)
    option_b_summary: str = ""


# ============================================================================
# Helper Functions: Text & Image Hashing
# ============================================================================

def normalize_text(text: str) -> str:
    """Normalize whitespace while preserving case and punctuation."""
    if not text:
        return ""
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def normalize_for_matching(text: str) -> str:
    """Lowercase alphanumeric normalization for fuzzy matching."""
    if not text:
        return ""
    cleaned = re.sub(r"[^a-z0-9\s\.\-]", " ", text.lower())
    return re.sub(r"\s+", " ", cleaned).strip()


def compute_dhash(image_bytes: bytes, hash_size: int = 8) -> Optional[int]:
    """
    Compute an 8x8 perceptual difference hash (dHash) using Pillow.
    Returns None if Pillow is unavailable or the image format cannot be decoded
    (e.g., WMF/EMF/SVG vector formats, which fall back to SHA-256).
    """
    if not HAS_PILLOW:
        return None
    try:
        with Image.open(io.BytesIO(image_bytes)) as img:
            gray = img.convert("L").resize(
                (hash_size + 1, hash_size), Image.Resampling.LANCZOS
            )
            if hasattr(gray, "get_flattened_data"):
                pixels = list(gray.get_flattened_data())
            else:
                pixels = list(gray.getdata())
            diff_bits = 0
            for row in range(hash_size):
                row_start = row * (hash_size + 1)
                for col in range(hash_size):
                    diff_bits <<= 1
                    if pixels[row_start + col] > pixels[row_start + col + 1]:
                        diff_bits |= 1
            return diff_bits
    except Exception:
        return None


def images_are_visually_same(img1: ImageItem, img2: ImageItem, max_hamming: int = 4) -> bool:
    """
    Return True if two images have identical bytes (SHA-256) or nearly identical
    perceptual dHash (so resized/re-saved copies of the same picture match).
    """
    if img1.sha256 == img2.sha256:
        return True
    if img1.dhash is not None and img2.dhash is not None:
        hamming = bin(img1.dhash ^ img2.dhash).count("1")
        return hamming <= max_hamming
    return False


def format_slide_list(slides: List[int], prefix: str = "Slide", plural_prefix: str = "Slides") -> str:
    """Format a list of slide numbers cleanly, e.g., 'Slide 5' or 'Slides 7, 14'."""
    unique_sorted = sorted(set(slides))
    if not unique_sorted:
        return ""
    nums = ", ".join(str(s) for s in unique_sorted)
    if len(unique_sorted) == 1:
        return f"{prefix} {nums}" if prefix else nums
    return f"{plural_prefix} {nums}" if plural_prefix else nums


def format_slide_nums_compact(slides: List[int]) -> str:
    """Format slide numbers as compact comma-separated string, e.g., '6, 7, 11, 18'."""
    return ", ".join(str(s) for s in sorted(set(slides)))


# ============================================================================
# Extraction Engine
# ============================================================================

def extract_run_formatting_sig(text_frame) -> str:
    """Build a compact signature of font names, sizes, bold, italic, and color."""
    parts = []
    try:
        for para in text_frame.paragraphs:
            p_align = str(para.alignment) if para.alignment else ""
            for run in para.runs:
                r_text = run.text.strip()
                if not r_text:
                    continue
                font = run.font
                f_name = font.name or ""
                f_size = str(font.size.pt) if getattr(font, "size", None) else ""
                f_bold = "B" if font.bold else ""
                f_italic = "I" if font.italic else ""
                f_under = "U" if font.underline else ""
                f_color = ""
                try:
                    if font.color and font.color.type is not None and hasattr(font.color, "rgb") and font.color.rgb:
                        f_color = str(font.color.rgb)
                except Exception:
                    pass
                parts.append(f"{p_align}:{f_name}:{f_size}:{f_bold}{f_italic}{f_under}:{f_color}")
    except Exception:
        pass
    return ";".join(parts)


def iter_flattened_shapes(shapes, offset_left: int = 0, offset_top: int = 0):
    """
    Recursively yield (shape, abs_left, abs_top, width, height) for all shapes,
    including shapes nested inside GroupShapes.
    """
    for shape in shapes:
        s_left = (getattr(shape, "left", 0) or 0) + offset_left
        s_top = (getattr(shape, "top", 0) or 0) + offset_top
        s_width = getattr(shape, "width", 0) or 0
        s_height = getattr(shape, "height", 0) or 0

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                yield from iter_flattened_shapes(shape.shapes, s_left, s_top)
                continue
            except Exception:
                pass
        yield shape, s_left, s_top, s_width, s_height


def extract_diagram_texts_from_slide(slide) -> List[str]:
    """Extract text from SmartArt / Diagram XML parts associated with a slide."""
    diagram_texts: List[str] = []
    try:
        for rel in slide.part.rels.values():
            if "diagramData" in rel.reltype:
                xml_bytes = rel.target_part.blob
                root = ET.fromstring(xml_bytes)
                texts = [
                    elem.text.strip()
                    for elem in root.iter()
                    if elem.tag.endswith("}t") and elem.text and elem.text.strip()
                ]
                if texts:
                    diagram_texts.append(" | ".join(texts))
    except Exception:
        pass
    return diagram_texts


def extract_presentation(
    pptx_path: Path,
    version_tag: str,
    run_dir: Path,
) -> PresentationData:
    """
    Parse a .pptx file, extract all images and slide content into run_dir,
    classify slides according to the training hierarchy, and return PresentationData.
    """
    prs = Presentation(str(pptx_path))
    slide_w = int(prs.slide_width or 9144000)
    slide_h = int(prs.slide_height or 6858000)

    images_dir = run_dir / f"{version_tag}_images"
    images_dir.mkdir(parents=True, exist_ok=True)

    deck_data = PresentationData(
        file_path=pptx_path,
        version_tag=version_tag,
        slide_width=slide_w,
        slide_height=slide_h,
        images_dir=images_dir,
    )

    for idx, slide in enumerate(prs.slides, start=1):
        slide_id = int(getattr(slide, "slide_id", idx))

        # Determine the title shape if present
        title_shape_id = None
        title_text = ""
        try:
            if slide.shapes.title and slide.shapes.title.has_text_frame:
                title_candidate = slide.shapes.title.text_frame.text.strip()
                if title_candidate:
                    title_shape_id = slide.shapes.title.shape_id
                    title_text = normalize_text(title_candidate)
        except Exception:
            pass

        text_boxes: List[TextBoxItem] = []
        images: List[ImageItem] = []
        tables: List[TableItem] = []
        charts: List[ChartItem] = []
        cliparts: List[ClipartItem] = []
        seen_image_rids: Set[str] = set()
        img_counter = 0

        for shape, s_left, s_top, s_width, s_height in iter_flattened_shapes(slide.shapes):
            norm_l = round(s_left / slide_w, 4) if slide_w else 0.0
            norm_t = round(s_top / slide_h, 4) if slide_h else 0.0
            norm_w = round(s_width / slide_w, 4) if slide_w else 0.0
            norm_h = round(s_height / slide_h, 4) if slide_h else 0.0
            pos_sig = f"{round(norm_l, 2)},{round(norm_t, 2)},{round(norm_w, 2)},{round(norm_h, 2)}"
            shape_name = getattr(shape, "name", "") or ""

            # 1. Check for Image (direct shape.image or embedded blip in XML)
            extracted_img = False
            try:
                if hasattr(shape, "image"):
                    img_obj = shape.image
                    img_bytes = img_obj.blob
                    ext = (img_obj.ext or "png").lower().lstrip(".")
                    img_counter += 1
                    img_filename = f"slide_{idx:02d}_img_{img_counter:02d}.{ext}"
                    img_out_path = images_dir / img_filename
                    img_out_path.write_bytes(img_bytes)

                    sha = hashlib.sha256(img_bytes).hexdigest()
                    dhash = compute_dhash(img_bytes)
                    images.append(
                        ImageItem(
                            index_on_slide=img_counter,
                            filename=img_filename,
                            rel_path=f"{images_dir.name}/{img_filename}",
                            ext=ext,
                            sha256=sha,
                            dhash=dhash,
                            byte_size=len(img_bytes),
                            norm_left=norm_l,
                            norm_top=norm_t,
                            norm_width=norm_w,
                            norm_height=norm_h,
                            format_signature=pos_sig,
                        )
                    )
                    extracted_img = True
                    for blip in shape._element.iter():
                        if blip.tag.endswith("}blip"):
                            for attr_k, attr_v in blip.attrib.items():
                                if attr_k.endswith("}embed"):
                                    seen_image_rids.add(attr_v)
            except Exception:
                pass

            if not extracted_img:
                try:
                    for blip in shape._element.iter():
                        if blip.tag.endswith("}blip"):
                            for attr_k, attr_v in blip.attrib.items():
                                if attr_k.endswith("}embed") and attr_v not in seen_image_rids:
                                    seen_image_rids.add(attr_v)
                                    rel_part = slide.part.related_parts.get(attr_v)
                                    if rel_part is not None:
                                        img_bytes = rel_part.blob
                                        ext = Path(str(rel_part.partname)).suffix.lstrip(".").lower() or "png"
                                        img_counter += 1
                                        img_filename = f"slide_{idx:02d}_img_{img_counter:02d}.{ext}"
                                        img_out_path = images_dir / img_filename
                                        img_out_path.write_bytes(img_bytes)

                                        sha = hashlib.sha256(img_bytes).hexdigest()
                                        dhash = compute_dhash(img_bytes)
                                        images.append(
                                            ImageItem(
                                                index_on_slide=img_counter,
                                                filename=img_filename,
                                                rel_path=f"{images_dir.name}/{img_filename}",
                                                ext=ext,
                                                sha256=sha,
                                                dhash=dhash,
                                                byte_size=len(img_bytes),
                                                norm_left=norm_l,
                                                norm_top=norm_t,
                                                norm_width=norm_w,
                                                norm_height=norm_h,
                                                format_signature=pos_sig,
                                            )
                                        )
                                        extracted_img = True
                except Exception:
                    pass

            # 2. Check for Table
            if getattr(shape, "has_table", False):
                try:
                    rows_data: List[List[str]] = []
                    for row in shape.table.rows:
                        rows_data.append([normalize_text(cell.text) for cell in row.cells])
                    tbl_text = "\n".join(" | ".join(r) for r in rows_data)
                    tables.append(
                        TableItem(
                            rows=rows_data,
                            normalized_text=normalize_text(tbl_text),
                            format_signature=f"table:{pos_sig}:{len(rows_data)}",
                        )
                    )
                except Exception:
                    pass

            # 3. Check for Chart
            if getattr(shape, "has_chart", False):
                try:
                    chart = shape.chart
                    c_type = str(getattr(chart, "chart_type", "Chart"))
                    c_title = ""
                    if getattr(chart, "has_title", False) and chart.chart_title.has_text_frame:
                        c_title = normalize_text(chart.chart_title.text_frame.text)
                    series_info: List[str] = []
                    for s in chart.series:
                        s_name = str(getattr(s, "name", "") or "")
                        try:
                            vals = [str(v) for v in (s.values or [])]
                        except Exception:
                            vals = []
                        series_info.append(f"{s_name}: [{', '.join(vals)}]")
                    c_norm = normalize_text(f"{c_type} | {c_title} | " + "; ".join(series_info))
                    charts.append(
                        ChartItem(
                            chart_type=c_type,
                            title=c_title,
                            series_data=series_info,
                            normalized_text=c_norm,
                            format_signature=f"chart:{pos_sig}:{c_type}",
                        )
                    )
                except Exception:
                    pass

            # 4. Check for Text Frame
            has_nonempty_text = False
            if getattr(shape, "has_text_frame", False):
                try:
                    raw_text = shape.text_frame.text
                    norm_txt = normalize_text(raw_text)
                    if norm_txt:
                        has_nonempty_text = True
                        run_sig = extract_run_formatting_sig(shape.text_frame)
                        is_title_box = (
                            title_shape_id is not None
                            and getattr(shape, "shape_id", None) == title_shape_id
                        )
                        text_boxes.append(
                            TextBoxItem(
                                shape_name=shape_name,
                                text=raw_text.strip(),
                                normalized_text=norm_txt,
                                norm_left=norm_l,
                                norm_top=norm_t,
                                norm_width=norm_w,
                                norm_height=norm_h,
                                format_signature=f"{pos_sig}:{run_sig}",
                                is_title=is_title_box,
                            )
                        )
                except Exception:
                    pass

            # 5. Check for Clipart / Vector AutoShapes / Freeforms
            if (
                not extracted_img
                and not getattr(shape, "has_table", False)
                and not getattr(shape, "has_chart", False)
            ):
                st = getattr(shape, "shape_type", None)
                if st in (
                    MSO_SHAPE_TYPE.AUTO_SHAPE,
                    MSO_SHAPE_TYPE.FREEFORM,
                    MSO_SHAPE_TYPE.CANVAS,
                    MSO_SHAPE_TYPE.IGX_GRAPHIC,
                ) and not getattr(shape, "is_placeholder", False):
                    if not has_nonempty_text or st in (MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.CANVAS, MSO_SHAPE_TYPE.IGX_GRAPHIC):
                        s_txt = normalize_text(shape.text_frame.text) if getattr(shape, "has_text_frame", False) else ""
                        cliparts.append(
                            ClipartItem(
                                shape_type=str(st),
                                shape_name=shape_name,
                                text=s_txt,
                                format_signature=f"clipart:{st}:{pos_sig}",
                            )
                        )

        # Fallback title selection if slide.shapes.title was empty
        if not title_text and text_boxes:
            upper_boxes = [tb for tb in text_boxes if tb.norm_top <= 0.30]
            if upper_boxes:
                top_box = min(upper_boxes, key=lambda b: (b.norm_top, b.norm_left))
                top_box.is_title = True
                title_text = top_box.normalized_text

        # Extract Speaker Notes
        notes_raw = ""
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_raw = slide.notes_slide.notes_text_frame.text or ""
        except Exception:
            pass

        diagram_texts = extract_diagram_texts_from_slide(slide)

        slide_data = SlideData(
            slide_number=idx,
            slide_id=slide_id,
            category=CAT_CONTENT,  # Refined in classify_deck_slides()
            title=title_text,
            normalized_title=normalize_for_matching(title_text),
            text_boxes=text_boxes,
            images=images,
            tables=tables,
            charts=charts,
            cliparts=cliparts,
            diagram_texts=diagram_texts,
            speaker_notes=notes_raw.strip(),
            normalized_notes=normalize_text(notes_raw),
        )
        deck_data.slides.append(slide_data)

    # Classify slides and populate reference catalog from References slide(s)
    classify_deck_slides(deck_data)
    deck_data.reference_catalog = build_reference_catalog(deck_data)

    return deck_data


# ============================================================================
# Slide Classification & On-Slide Reference Detection
# ============================================================================

def classify_deck_slides(deck: PresentationData) -> None:
    """
    Classify each slide into the standard training deck categories:
    - Title Slide
    - Course Terminal Objectives
    - Module Terminal Objectives
    - Lesson Learning Objective
    - Section Enabling Objective
    - References
    - Knowledge Check
    - Content
    """
    terminal_obj_slides: List[SlideData] = []

    for slide in deck.slides:
        heading_lower = slide.title.lower()
        upper_text_lower = "\n".join(
            tb.normalized_text.lower() for tb in slide.text_boxes if tb.norm_top <= 0.35
        )
        all_text_lower = "\n".join(tb.normalized_text.lower() for tb in slide.text_boxes)

        # 1. Knowledge Check (check slide title and any upper-slide header text box)
        if "knowledge check" in heading_lower or "knowledge check" in upper_text_lower:
            slide.category = CAT_KNOWLEDGE_CHECK
            continue

        # 2. References Slide
        heading_clean = re.sub(r"[^a-z\s]", "", heading_lower).strip()
        upper_lines_clean = [
            re.sub(r"[^a-z\s]", "", ln).strip()
            for ln in upper_text_lower.splitlines()
            if ln.strip()
        ]
        ref_headers = {"reference", "references", "course references", "training references", "reference list"}
        if heading_clean in ref_headers or any(ln in ref_headers for ln in upper_lines_clean):
            slide.category = CAT_REFERENCES
            continue

        # 3. Objectives Slides (spelled out: Terminal Objective(s), Learning Objective(s), Enabling Objective(s))
        if "terminal objective" in heading_lower or "terminal objective" in all_text_lower:
            if "course" in heading_lower or "course terminal objective" in all_text_lower:
                slide.category = CAT_COURSE_TLO
            elif "module" in heading_lower or "module terminal objective" in all_text_lower:
                slide.category = CAT_MODULE_TLO
            else:
                slide.category = CAT_GENERIC_TLO
                terminal_obj_slides.append(slide)
            continue

        if "enabling objective" in heading_lower or "enabling objective" in all_text_lower:
            slide.category = CAT_SECTION_EO
            continue

        if "learning objective" in heading_lower or "learning objective" in all_text_lower:
            slide.category = CAT_LESSON_LO
            continue

        # 4. Title Slide (Slide 1)
        if slide.slide_number == 1:
            slide.category = CAT_TITLE
            continue

        slide.category = CAT_CONTENT

    # If there are two generic "Terminal Objectives" slides near the beginning that didn't
    # explicitly say "Course" vs "Module", the training hierarchy places Course Terminal
    # Objectives first and Module Terminal Objectives second.
    if len(terminal_obj_slides) >= 2:
        terminal_obj_slides[0].category = CAT_COURSE_TLO
        terminal_obj_slides[1].category = CAT_MODULE_TLO


def build_reference_catalog(deck: PresentationData) -> List[str]:
    """
    Extract individual reference lines/items from the main References slide(s).
    """
    catalog: List[str] = []
    for slide in deck.slides:
        if slide.category == CAT_REFERENCES:
            for tb in slide.text_boxes:
                if tb.is_title:
                    continue
                for line in tb.normalized_text.splitlines():
                    cleaned = re.sub(r"^[\•\-\*\d\.\)\s]+", "", line).strip()
                    if len(cleaned) >= 3 and cleaned.lower() not in ("reference", "references"):
                        catalog.append(cleaned)
    return catalog


def tokenize_reference_text(text: str) -> Set[str]:
    """Extract meaningful reference tokens (alphanumeric codes, document numbers, keywords)."""
    tokens = set(re.findall(r"[a-z0-9][a-z0-9\-\./]{2,}", text.lower()))
    return {t.strip(".-/") for t in tokens if t.strip(".-/") not in STOPWORDS and len(t.strip(".-/")) >= 3}


def partition_slide_text_boxes(
    deck: PresentationData,
    combined_reference_catalog: List[str],
) -> None:
    """
    Separate each slide's text boxes into:
    - Slide Title
    - On-Slide Reference Text Boxes (matched against the References slide catalog and bottom-left position)
    - Body Content Text Boxes
    """
    catalog_phrases: List[str] = []
    catalog_token_sets: List[Set[str]] = []
    all_catalog_tokens: Set[str] = set()

    for entry in combined_reference_catalog:
        norm_entry = normalize_for_matching(entry)
        if len(norm_entry) >= 4:
            catalog_phrases.append(norm_entry)
        for part in re.split(r"[,;:\-\(\)]", entry):
            norm_part = normalize_for_matching(part)
            if len(norm_part) >= 5:
                catalog_phrases.append(norm_part)
        t_set = tokenize_reference_text(entry)
        if t_set:
            catalog_token_sets.append(t_set)
            all_catalog_tokens.update(t_set)

    for slide in deck.slides:
        slide.body_text_items = []
        slide.on_slide_references = []

        if slide.category == CAT_REFERENCES:
            for tb in slide.text_boxes:
                if not tb.is_title:
                    slide.body_text_items.append(tb.normalized_text)
            continue

        for tb in slide.text_boxes:
            if tb.is_title:
                continue

            box_text = tb.normalized_text
            box_norm = normalize_for_matching(box_text)
            if not box_norm:
                continue

            is_ref_box = False

            # Check 1: Explicit "Ref:" / "Reference:" / "References:" prefix
            if re.match(r"^(refs?|references?|sources?)\s*[:\-]", box_text.strip(), re.IGNORECASE):
                is_ref_box = True

            # Check 2: Position + content match against the main References slide catalog
            is_bottom_area = (tb.norm_top >= 0.55) or ((tb.norm_top + tb.norm_height) >= 0.72)
            is_bottom_left_favor = tb.norm_left <= 0.65

            if not is_ref_box and catalog_phrases:
                box_lines = [normalize_for_matching(ln) for ln in box_text.splitlines() if ln.strip()]
                matched_lines = 0
                for ln in box_lines:
                    if len(ln) < 4:
                        continue
                    for phrase in catalog_phrases:
                        if (len(ln) >= 5 and ln in phrase) or (len(phrase) >= 5 and phrase in ln):
                            matched_lines += 1
                            break
                        else:
                            if is_bottom_area and len(ln) >= 6 and len(phrase) >= 6:
                                if difflib.SequenceMatcher(None, ln, phrase).ratio() >= 0.72:
                                    matched_lines += 1
                                    break

                if box_lines and (matched_lines / len(box_lines)) >= 0.5:
                    if len(box_lines) <= 6 or is_bottom_area:
                        is_ref_box = True

                if not is_ref_box and is_bottom_area and is_bottom_left_favor:
                    box_tokens = tokenize_reference_text(box_text)
                    if box_tokens:
                        overlap = box_tokens & all_catalog_tokens
                        overlap_ratio = len(overlap) / len(box_tokens)
                        if overlap_ratio >= 0.50:
                            is_ref_box = True
                        else:
                            for ref_tokens in catalog_token_sets:
                                if ref_tokens and len(box_tokens & ref_tokens) / min(len(box_tokens), len(ref_tokens)) >= 0.60:
                                    is_ref_box = True
                                    break

            tb.is_reference = is_ref_box
            if is_ref_box:
                slide.on_slide_references.append(box_text)
            else:
                slide.body_text_items.append(box_text)


# ============================================================================
# Writing Extracted Representation Files
# ============================================================================

def write_presentation_representation(deck: PresentationData, run_dir: Path) -> Path:
    """
    Write a human-readable text representation of the presentation linking all
    extracted images, slide categories, body text, tables, charts, on-slide
    references, and speaker notes.
    """
    out_path = run_dir / f"{deck.version_tag}_extracted.txt"
    lines: List[str] = []
    lines.append("=" * 80)
    lines.append(f"PRESENTATION: {deck.file_path.name}")
    lines.append(f"Version Tag : {deck.version_tag}")
    lines.append(f"Source Path : {deck.file_path}")
    lines.append(f"Total Slides: {len(deck.slides)}")
    lines.append(f"Images Dir  : {deck.images_dir.name if deck.images_dir else 'N/A'}")
    lines.append("=" * 80)
    lines.append("")

    if deck.reference_catalog:
        lines.append("MAIN REFERENCES CATALOG (from References Slide):")
        for r_idx, ref_item in enumerate(deck.reference_catalog, start=1):
            lines.append(f"  [{r_idx}] {ref_item}")
        lines.append("")

    for slide in deck.slides:
        lines.append("-" * 80)
        lines.append(
            f"SLIDE {slide.slide_number} (ID: {slide.slide_id}) | "
            f"CATEGORY: {slide.category}"
        )
        lines.append(f"TITLE: {slide.title or '(No Title)'}")
        lines.append("-" * 80)

        if slide.combined_body_text:
            lines.append("[Slide Content]")
            for ln in slide.combined_body_text.splitlines():
                lines.append(f"  {ln}")
            lines.append("")

        if slide.tables:
            lines.append(f"[Tables ({len(slide.tables)})]")
            for t_idx, tbl in enumerate(slide.tables, start=1):
                lines.append(f"  Table {t_idx}:")
                for row in tbl.rows:
                    lines.append("    | " + " | ".join(row) + " |")
            lines.append("")

        if slide.charts:
            lines.append(f"[Charts ({len(slide.charts)})]")
            for c_idx, cht in enumerate(slide.charts, start=1):
                lines.append(f"  Chart {c_idx} ({cht.chart_type}): {cht.title}")
                for s_str in cht.series_data:
                    lines.append(f"    - {s_str}")
            lines.append("")

        if slide.cliparts:
            lines.append(f"[Clipart / Graphic Shapes ({len(slide.cliparts)})]")
            for g_idx, clp in enumerate(slide.cliparts, start=1):
                label = f" ({clp.text})" if clp.text else ""
                lines.append(f"  Graphic {g_idx}: {clp.shape_name}{label}")
            lines.append("")

        if slide.images:
            lines.append(f"[Images ({len(slide.images)})]")
            for img in slide.images:
                lines.append(
                    f"  - Image {img.index_on_slide}: {img.rel_path} "
                    f"(SHA256: {img.sha256[:12]}..., Size: {img.byte_size} bytes)"
                )
            lines.append("")

        if slide.on_slide_references:
            lines.append("[On-Slide References]")
            for ref in slide.on_slide_references:
                for ref_ln in ref.splitlines():
                    lines.append(f"  * {ref_ln}")
            lines.append("")

        if slide.normalized_notes:
            lines.append("[Speaker Notes]")
            for n_ln in slide.normalized_notes.splitlines():
                lines.append(f"  {n_ln}")
            lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    deck.extracted_text_path = out_path
    return out_path


# ============================================================================
# Slide Alignment & Comparison Engine
# ============================================================================

def slide_similarity(s1: SlideData, s2: SlideData) -> float:
    """Compute similarity score [0.0, 1.0] between two slides."""
    if s1.category != s2.category:
        if s1.category != CAT_CONTENT or s2.category != CAT_CONTENT:
            return 0.0

    title_sim = difflib.SequenceMatcher(None, s1.normalized_title, s2.normalized_title).ratio()
    body1 = normalize_for_matching(s1.combined_body_text)
    body2 = normalize_for_matching(s2.combined_body_text)
    body_sim = difflib.SequenceMatcher(None, body1, body2).ratio() if (body1 or body2) else 1.0

    if s1.normalized_title and s2.normalized_title:
        return 0.55 * title_sim + 0.45 * body_sim
    return body_sim


def align_slides(
    old_slides: List[SlideData],
    new_slides: List[SlideData],
) -> Tuple[List[Tuple[SlideData, SlideData]], List[SlideData], List[SlideData]]:
    """
    Align slides between old_slides and new_slides so inserted or deleted slides
    do not shift and falsely flag subsequent slides.
    Returns (matched_pairs, added_slides, deleted_slides).
    """
    matched_old_indices: Set[int] = set()
    matched_new_indices: Set[int] = set()
    matched_pairs: List[Tuple[SlideData, SlideData]] = []

    # Pass 1: Match by PowerPoint's internal slide_id when categories match or similarity is reasonable
    new_by_id: Dict[int, int] = {}
    for n_idx, ns in enumerate(new_slides):
        if ns.slide_id not in new_by_id:
            new_by_id[ns.slide_id] = n_idx

    for o_idx, os_item in enumerate(old_slides):
        if os_item.slide_id in new_by_id:
            n_idx = new_by_id[os_item.slide_id]
            ns_item = new_slides[n_idx]
            sim = slide_similarity(os_item, ns_item)
            if os_item.category == ns_item.category and (sim >= 0.25 or os_item.normalized_title == ns_item.normalized_title):
                matched_old_indices.add(o_idx)
                matched_new_indices.add(n_idx)
                matched_pairs.append((os_item, ns_item))

    # Pass 2: Match unique singleton category slides (Title, Course TLO, Module TLO, Lesson LO, Section EO, References)
    singleton_cats = {
        CAT_TITLE,
        CAT_COURSE_TLO,
        CAT_MODULE_TLO,
        CAT_LESSON_LO,
        CAT_SECTION_EO,
        CAT_REFERENCES,
    }
    for cat in singleton_cats:
        unmatched_old = [i for i, s in enumerate(old_slides) if i not in matched_old_indices and s.category == cat]
        unmatched_new = [j for j, s in enumerate(new_slides) if j not in matched_new_indices and s.category == cat]
        if len(unmatched_old) == 1 and len(unmatched_new) == 1:
            o_idx = unmatched_old[0]
            n_idx = unmatched_new[0]
            matched_old_indices.add(o_idx)
            matched_new_indices.add(n_idx)
            matched_pairs.append((old_slides[o_idx], new_slides[n_idx]))

    # Pass 3: Greedy best-similarity match for remaining unmatched slides
    candidates: List[Tuple[float, int, int]] = []
    for o_idx, os_item in enumerate(old_slides):
        if o_idx in matched_old_indices:
            continue
        for n_idx, ns_item in enumerate(new_slides):
            if n_idx in matched_new_indices:
                continue
            sim = slide_similarity(os_item, ns_item)
            pos_bonus = 0.05 * (1.0 - abs(o_idx - n_idx) / max(len(old_slides), len(new_slides), 1))
            if sim >= 0.50:
                candidates.append((sim + pos_bonus, o_idx, n_idx))

    candidates.sort(reverse=True, key=lambda x: x[0])
    for _, o_idx, n_idx in candidates:
        if o_idx not in matched_old_indices and n_idx not in matched_new_indices:
            matched_old_indices.add(o_idx)
            matched_new_indices.add(n_idx)
            matched_pairs.append((old_slides[o_idx], new_slides[n_idx]))

    matched_pairs.sort(key=lambda pair: pair[1].slide_number)
    added_slides = [s for j, s in enumerate(new_slides) if j not in matched_new_indices]
    deleted_slides = [s for i, s in enumerate(old_slides) if i not in matched_old_indices]

    return matched_pairs, added_slides, deleted_slides


def compare_slide_images(
    old_images: List[ImageItem],
    new_images: List[ImageItem],
) -> Tuple[int, int, int, bool]:
    """
    Compare images on a slide using visual content hashing (SHA256 + perceptual dHash).
    Moving or slightly resizing an image on the slide does NOT count as an image change;
    instead, if the image's position/size signature changed, it sets minor_format=True.
    Returns (added_count, deleted_count, replaced_count, minor_format_changed).
    """
    used_old: Set[int] = set()
    used_new: Set[int] = set()
    minor_format = False

    for n_idx, n_img in enumerate(new_images):
        for o_idx, o_img in enumerate(old_images):
            if o_idx in used_old:
                continue
            if images_are_visually_same(o_img, n_img):
                used_old.add(o_idx)
                used_new.add(n_idx)
                if o_img.format_signature != n_img.format_signature:
                    minor_format = True
                break

    unmatched_old = len(old_images) - len(used_old)
    unmatched_new = len(new_images) - len(used_new)
    replaced = min(unmatched_old, unmatched_new)
    added = unmatched_new - replaced
    deleted = unmatched_old - replaced

    return added, deleted, replaced, minor_format


def compare_presentations(
    old_deck: PresentationData,
    new_deck: PresentationData,
) -> DeckComparisonResult:
    """
    Compare two consecutive presentation versions and populate DeckComparisonResult.
    """
    combined_refs = list(dict.fromkeys(old_deck.reference_catalog + new_deck.reference_catalog))
    partition_slide_text_boxes(old_deck, combined_refs)
    partition_slide_text_boxes(new_deck, combined_refs)

    result = DeckComparisonResult(old_deck=old_deck, new_deck=new_deck)

    matched_pairs, added_slides, deleted_slides = align_slides(old_deck.slides, new_deck.slides)
    result.added_slides = added_slides
    result.deleted_slides = deleted_slides

    # Check if any matched slides were reordered relative to each other
    last_old_num = 0
    for os_item, ns_item in matched_pairs:
        if os_item.slide_number < last_old_num:
            result.reordered_slides.append((os_item.slide_number, ns_item.slide_number))
        last_old_num = max(last_old_num, os_item.slide_number)

    # Record added slides in category-specific trackers as well
    for ns in added_slides:
        s_num = ns.slide_number
        if ns.category == CAT_KNOWLEDGE_CHECK:
            result.kc_added_slides.append(s_num)
        elif ns.category in OBJECTIVE_CATEGORIES:
            result.objectives_changed.append((ns.category, s_num))
        elif ns.category == CAT_REFERENCES:
            result.main_ref_slides_changed.append(s_num)
        if ns.images:
            result.image_changed_slides.append(s_num)
            result.images_added_slides.append(s_num)
        if ns.tables:
            result.tables_changed_slides.append(s_num)
        if ns.charts:
            result.charts_changed_slides.append(s_num)
        if ns.cliparts:
            result.clipart_changed_slides.append(s_num)
        if ns.on_slide_references:
            result.on_slide_refs_changed.append(s_num)

        result.slide_details.append(
            SlideDiffDetail(
                old_slide_num=None,
                new_slide_num=s_num,
                category=ns.category,
                title=ns.title,
                status="added",
                change_notes=[f"New {ns.category} slide added ('{ns.title or 'Untitled'}')"],
            )
        )

    for os_item in deleted_slides:
        if os_item.category == CAT_KNOWLEDGE_CHECK:
            result.kc_deleted_old_slides.append(os_item.slide_number)
        result.slide_details.append(
            SlideDiffDetail(
                old_slide_num=os_item.slide_number,
                new_slide_num=None,
                category=os_item.category,
                title=os_item.title,
                status="deleted",
                change_notes=[f"Removed {os_item.category} slide (old Slide {os_item.slide_number}: '{os_item.title or 'Untitled'}')"],
            )
        )

    # Compare each matched slide pair
    for os_item, ns_item in matched_pairs:
        s_num = ns_item.slide_number
        detail = SlideDiffDetail(
            old_slide_num=os_item.slide_number,
            new_slide_num=s_num,
            category=ns_item.category,
            title=ns_item.title or os_item.title,
            status="unchanged",
        )

        # 1. Title & Body Text
        title_changed = os_item.title != ns_item.title
        body_changed = os_item.normalized_body_text != ns_item.normalized_body_text
        text_changed = title_changed or body_changed

        if text_changed:
            detail.text_changed = True
            if ns_item.category in OBJECTIVE_CATEGORIES:
                result.objectives_changed.append((ns_item.category, s_num))
                detail.change_notes.append(f"{ns_item.category} text updated")
            elif ns_item.category == CAT_REFERENCES:
                result.main_ref_slides_changed.append(s_num)
                detail.change_notes.append("Main References slide list updated")
            elif ns_item.category == CAT_KNOWLEDGE_CHECK:
                result.kc_updated_slides.append(s_num)
                detail.change_notes.append("Knowledge Check question/content text updated")
            else:
                result.content_text_changed_slides.append(s_num)
                detail.change_notes.append("Slide text updated")

        # 2. On-Slide References
        if os_item.normalized_refs_text != ns_item.normalized_refs_text:
            detail.refs_changed = True
            result.on_slide_refs_changed.append(s_num)
            detail.change_notes.append(
                f"On-slide reference updated: '{os_item.normalized_refs_text or 'None'}' -> '{ns_item.normalized_refs_text or 'None'}'"
            )

        # 3. Speaker Notes
        if os_item.normalized_notes != ns_item.normalized_notes:
            detail.notes_changed = True
            result.speaker_notes_changed_slides.append(s_num)
            detail.change_notes.append("Speaker notes updated")
            if ns_item.category == CAT_KNOWLEDGE_CHECK and s_num not in result.kc_updated_slides:
                result.kc_updated_slides.append(s_num)

        # 4. Images
        img_add, img_del, img_rep, img_fmt = compare_slide_images(os_item.images, ns_item.images)
        detail.images_added = img_add
        detail.images_deleted = img_del
        detail.images_replaced = img_rep
        if img_add or img_del or img_rep:
            result.image_changed_slides.append(s_num)
            if img_add:
                result.images_added_slides.append(s_num)
            if img_del:
                result.images_deleted_slides.append(s_num)
            if img_rep:
                result.images_replaced_slides.append(s_num)
            img_parts = []
            if img_add:
                img_parts.append(f"{img_add} added")
            if img_rep:
                img_parts.append(f"{img_rep} replaced")
            if img_del:
                img_parts.append(f"{img_del} deleted")
            detail.change_notes.append(f"Images changed ({', '.join(img_parts)})")

        # 5. Tables
        old_tbl_texts = [t.normalized_text for t in os_item.tables]
        new_tbl_texts = [t.normalized_text for t in ns_item.tables]
        if old_tbl_texts != new_tbl_texts:
            detail.tables_changed = True
            result.tables_changed_slides.append(s_num)
            detail.change_notes.append(f"Tables changed ({len(os_item.tables)} -> {len(ns_item.tables)})")

        # 6. Charts
        old_cht_texts = [c.normalized_text for c in os_item.charts]
        new_cht_texts = [c.normalized_text for c in ns_item.charts]
        if old_cht_texts != new_cht_texts:
            detail.charts_changed = True
            result.charts_changed_slides.append(s_num)
            detail.change_notes.append(f"Charts changed ({len(os_item.charts)} -> {len(ns_item.charts)})")

        # 7. Clipart / Graphic Shapes
        old_clp_types = [(c.shape_type, c.text) for c in os_item.cliparts]
        new_clp_types = [(c.shape_type, c.text) for c in ns_item.cliparts]
        if old_clp_types != new_clp_types:
            detail.clipart_changed = True
            result.clipart_changed_slides.append(s_num)
            detail.change_notes.append(f"Clipart/shapes changed ({len(os_item.cliparts)} -> {len(ns_item.cliparts)})")

        # 8. Minor Formatting Changes
        has_content_diff = bool(detail.change_notes)
        if img_fmt or (not has_content_diff and os_item.full_format_signature != ns_item.full_format_signature):
            detail.minor_formatting = True
            result.minor_formatting_slides.append(s_num)
            if not has_content_diff:
                detail.status = "formatting_only"
                detail.change_notes.append("Minor formatting/layout adjustments")

        if has_content_diff:
            detail.status = "modified"

        result.slide_details.append(detail)

    result.slide_details.sort(key=lambda d: (d.new_slide_num or 9999, d.old_slide_num or 9999))
    result.option_b_summary = build_option_b_summary(result)
    return result


# ============================================================================
# Option B Summary Builder (3-4 Lines for Single Excel Cell)
# ============================================================================

def build_option_b_summary(res: DeckComparisonResult) -> str:
    """
    Generate the compact 3-4 line Option B summary designed to be pasted
    into a single Excel cell.

    Example format:
    Slides: 2 added (8, 15), 1 deleted (old 12) | Images changed: Slides 7, 14
    Objectives/Refs: Enabling Objective (Slide 5) & References (Slides 3, 8, 14) updated
    Knowledge Checks: Updated on Slides 10, 22
    Content/Notes: Slide text changed (6, 7, 11, 18); Speaker notes changed (6, 9, 10, 18, 22)
    Minor formatting changes
    """
    lines: List[str] = []

    # --- Line 1: Slides Added/Deleted + Images / Tables / Charts / Clipart ---
    slide_parts: List[str] = []
    if res.added_slides:
        add_nums = format_slide_nums_compact([s.slide_number for s in res.added_slides])
        slide_parts.append(f"{len(res.added_slides)} added ({add_nums})")
    if res.deleted_slides:
        del_nums = format_slide_nums_compact([s.slide_number for s in res.deleted_slides])
        slide_parts.append(f"{len(res.deleted_slides)} deleted (old {del_nums})")
    if not slide_parts:
        slide_parts.append("None added/deleted")

    line1_segments = [f"Slides: {', '.join(slide_parts)}"]

    if res.image_changed_slides:
        line1_segments.append(f"Images changed: {format_slide_list(res.image_changed_slides)}")
    else:
        line1_segments.append("Images changed: None")

    visual_extras: List[str] = []
    if res.tables_changed_slides:
        visual_extras.append(f"Tables ({format_slide_nums_compact(res.tables_changed_slides)})")
    if res.charts_changed_slides:
        visual_extras.append(f"Charts ({format_slide_nums_compact(res.charts_changed_slides)})")
    if res.clipart_changed_slides:
        visual_extras.append(f"Clipart ({format_slide_nums_compact(res.clipart_changed_slides)})")
    if visual_extras:
        line1_segments.append(f"Graphics/Tables: {', '.join(visual_extras)}")

    lines.append(" | ".join(line1_segments))

    # --- Line 2: Objectives & References ---
    obj_ref_parts: List[str] = []
    if res.objectives_changed:
        by_cat: Dict[str, List[int]] = {}
        for cat, s_num in res.objectives_changed:
            short_cat = cat.replace("Objectives", "Objective")
            by_cat.setdefault(short_cat, []).append(s_num)
        obj_strs = [f"{cat_name} ({format_slide_list(s_list)})" for cat_name, s_list in by_cat.items()]
        obj_ref_parts.append(", ".join(obj_strs))

    all_ref_slides = sorted(set(res.main_ref_slides_changed + res.on_slide_refs_changed))
    if all_ref_slides:
        obj_ref_parts.append(f"References ({format_slide_list(all_ref_slides)})")

    if obj_ref_parts:
        lines.append(f"Objectives/Refs: {' & '.join(obj_ref_parts)} updated")
    else:
        lines.append("Objectives/Refs: No changes")

    # --- Line 3: Knowledge Checks ---
    kc_parts: List[str] = []
    modified_only_kcs = [s for s in res.kc_updated_slides if s not in res.kc_added_slides]
    if modified_only_kcs:
        kc_parts.append(f"Updated on {format_slide_list(modified_only_kcs)}")
    if res.kc_added_slides:
        kc_parts.append(f"Added on {format_slide_list(res.kc_added_slides)}")
    if res.kc_deleted_old_slides:
        kc_parts.append(f"Deleted (old {format_slide_nums_compact(res.kc_deleted_old_slides)})")

    if kc_parts:
        lines.append(f"Knowledge Checks: {'; '.join(kc_parts)}")
    else:
        lines.append("Knowledge Checks: No changes")

    # --- Line 4: Content Text & Speaker Notes ---
    content_parts: List[str] = []
    if res.content_text_changed_slides:
        content_parts.append(f"Slide text changed ({format_slide_nums_compact(res.content_text_changed_slides)})")
    if res.speaker_notes_changed_slides:
        content_parts.append(f"Speaker notes changed ({format_slide_nums_compact(res.speaker_notes_changed_slides)})")

    if content_parts:
        lines.append(f"Content/Notes: {'; '.join(content_parts)}")
    else:
        lines.append("Content/Notes: No changes")

    # --- Optional Line 5: Minor Formatting Changes ---
    if res.minor_formatting_slides or res.reordered_slides:
        lines.append("Minor formatting changes")

    return "\n".join(lines)


# ============================================================================
# Report Writers (summary.txt & detailed_diff.txt)
# ============================================================================

def write_summary_and_details(
    comparisons: List[DeckComparisonResult],
    run_dir: Path,
) -> Tuple[Path, Path]:
    """
    Write summary.txt (containing the copy-paste ready Option B blocks for each
    version pair) and detailed_diff.txt (for slide-by-slide auditing).
    """
    summary_path = run_dir / "summary.txt"
    detailed_path = run_dir / "detailed_diff.txt"

    summary_lines: List[str] = []
    detailed_lines: List[str] = []

    for idx, comp in enumerate(comparisons, start=1):
        old_name = comp.old_deck.file_path.name
        new_name = comp.new_deck.file_path.name
        header = f"=== Comparison {idx}: {old_name} -> {new_name} ==="

        summary_lines.append(header)
        summary_lines.append(comp.option_b_summary)
        summary_lines.append("")

        detailed_lines.append("=" * 80)
        detailed_lines.append(header)
        detailed_lines.append("=" * 80)
        detailed_lines.append("HIGH-LEVEL SUMMARY (Option B):")
        detailed_lines.append(comp.option_b_summary)
        detailed_lines.append("")
        detailed_lines.append("SLIDE-BY-SLIDE BREAKDOWN:")

        any_detail = False
        for d in comp.slide_details:
            if d.status == "unchanged":
                continue
            any_detail = True
            if d.status == "added":
                loc = f"New Slide {d.new_slide_num}"
            elif d.status == "deleted":
                loc = f"Old Slide {d.old_slide_num} (Deleted)"
            elif d.old_slide_num != d.new_slide_num:
                loc = f"Slide {d.new_slide_num} (was Slide {d.old_slide_num})"
            else:
                loc = f"Slide {d.new_slide_num}"

            detailed_lines.append(f"  - [{loc} | {d.category}] {d.title or '(Untitled)'}:")
            for note in d.change_notes:
                detailed_lines.append(f"      * {note}")

        if not any_detail:
            detailed_lines.append("  No differences detected between these versions.")
        detailed_lines.append("")

    summary_path.write_text("\n".join(summary_lines).strip() + "\n", encoding="utf-8")
    detailed_path.write_text("\n".join(detailed_lines).strip() + "\n", encoding="utf-8")
    return summary_path, detailed_path


# ============================================================================
# Interactive File Picker Workflow (Tkinter)
# ============================================================================

def pick_pptx_files_interactively() -> List[Path]:
    """
    Open a step-by-step file picker allowing the user to select:
      1. Initial .pptx file (Version 1)
      2. Next .pptx file (Version 2, opening in the same folder)
      3. Prompt whether to add another version (Version 3, 4, ...) or finish.
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected_files: List[Path] = []
    last_dir: str = str(Path.cwd())

    # 1. Pick First File
    first_file = filedialog.askopenfilename(
        parent=root,
        title="Select Version 1 (Initial PPTX File)",
        initialdir=last_dir,
        filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")],
    )
    if not first_file:
        root.destroy()
        return []

    first_path = Path(first_file).resolve()
    selected_files.append(first_path)
    last_dir = str(first_path.parent)

    # 2. Pick Second File
    second_file = filedialog.askopenfilename(
        parent=root,
        title=f"Select Version 2 (Next PPTX after {first_path.name})",
        initialdir=last_dir,
        filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")],
    )
    if not second_file:
        root.destroy()
        return []

    second_path = Path(second_file).resolve()
    selected_files.append(second_path)
    last_dir = str(second_path.parent)

    # 3. Ask if user wants to keep adding versions (Version 3, Version 4, ...)
    while True:
        file_list_str = "\n".join(f"  {i}. {p.name}" for i, p in enumerate(selected_files, start=1))
        add_more = messagebox.askyesno(
            title="Add Another Version?",
            message=(
                f"Selected {len(selected_files)} versions so far:\n"
                f"{file_list_str}\n\n"
                f"Do you want to select another version (Version {len(selected_files) + 1}) "
                f"for this training?\n\n"
                f"Click 'Yes' to pick the next file, or 'No' to finish and run the comparison."
            ),
            parent=root,
        )
        if not add_more:
            break

        next_num = len(selected_files) + 1
        prev_name = selected_files[-1].name
        next_file = filedialog.askopenfilename(
            parent=root,
            title=f"Select Version {next_num} (Next PPTX after {prev_name})",
            initialdir=last_dir,
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")],
        )
        if not next_file:
            break

        next_path = Path(next_file).resolve()
        selected_files.append(next_path)
        last_dir = str(next_path.parent)

    root.destroy()
    return selected_files


# ============================================================================
# Main Runner
# ============================================================================

def sanitize_name(name: str) -> str:
    """Create a filesystem-safe folder/tag name."""
    cleaned = re.sub(r"[^A-Za-z0-9_\-]+", "_", name).strip("_")
    return cleaned or "presentation"


def run_pipeline(
    pptx_paths: List[Path],
    base_output_dir: Optional[Path] = None,
    open_summary: bool = True,
) -> Tuple[Path, Path, Path]:
    """
    Execute extraction and comparison across an ordered list of 2+ .pptx files.
    Returns (run_dir, summary_path, detailed_path).
    """
    if len(pptx_paths) < 2:
        raise ValueError("At least two .pptx files are required for comparison.")

    for p in pptx_paths:
        if not p.exists():
            raise FileNotFoundError(f"PPTX file not found: {p}")

    script_dir = Path(__file__).resolve().parent
    output_root = base_output_dir if base_output_dir is not None else (script_dir / "output")
    output_root.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base_stem = sanitize_name(pptx_paths[0].stem)
    run_dir = output_root / f"{base_stem}_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nWorking directory created: {run_dir}")

    # Step 1: Extract all presentations
    extracted_decks: List[PresentationData] = []
    for idx, p in enumerate(pptx_paths, start=1):
        v_tag = f"v{idx}_{sanitize_name(p.stem)}"
        print(f"  [{idx}/{len(pptx_paths)}] Extracting {p.name} ...")
        deck = extract_presentation(p, v_tag, run_dir)
        extracted_decks.append(deck)

    # Step 2: Compare each consecutive pair (v1 -> v2, v2 -> v3, ...)
    comparisons: List[DeckComparisonResult] = []
    for i in range(len(extracted_decks) - 1):
        old_deck = extracted_decks[i]
        new_deck = extracted_decks[i + 1]
        print(f"  Comparing: {old_deck.file_path.name} -> {new_deck.file_path.name} ...")
        comp = compare_presentations(old_deck, new_deck)
        comparisons.append(comp)

    # Step 3: Write the per-deck text representations (after on-slide references have been tagged)
    for deck in extracted_decks:
        write_presentation_representation(deck, run_dir)

    # Step 4: Write summary.txt and detailed_diff.txt
    summary_path, detailed_path = write_summary_and_details(comparisons, run_dir)

    print("\n" + "=" * 80)
    print("COMPARISON SUMMARY (Ready to copy-paste into Excel):")
    print("=" * 80)
    print(summary_path.read_text(encoding="utf-8"))
    print("=" * 80)
    print(f"Summary file saved to : {summary_path}")
    print(f"Detailed diff saved to: {detailed_path}")

    if open_summary and sys.platform.startswith("win"):
        try:
            os.startfile(str(summary_path))
        except Exception:
            pass

    return run_dir, summary_path, detailed_path


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compare two or more versions of a training PowerPoint (.pptx) presentation "
            "and generate a high-level summary of differences for Excel."
        )
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="Optional paths to 2 or more .pptx files in version order (if omitted, opens file picker).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Custom output directory (defaults to ./output next to pptx_diff.py).",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Do not automatically open summary.txt in the default text editor when finished.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    if args.files:
        if len(args.files) < 2:
            print("Error: Please provide at least 2 .pptx files when using command-line arguments, or run with no arguments to use the file picker.", file=sys.stderr)
            return 1
        selected_paths = [Path(f).resolve() for f in args.files]
    else:
        print("Opening file picker to select PowerPoint (.pptx) versions...")
        selected_paths = pick_pptx_files_interactively()
        if len(selected_paths) < 2:
            print("Fewer than 2 files selected. Exiting.")
            return 0

    try:
        run_pipeline(
            pptx_paths=selected_paths,
            base_output_dir=args.output_dir,
            open_summary=not args.no_open,
        )
    except Exception as exc:
        print(f"Error comparing presentations: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
