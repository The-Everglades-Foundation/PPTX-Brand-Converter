#!/usr/bin/env python3
"""Inspect a corporate PPTX/POTX template → brand_tokens.json."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from helpers import (  # noqa: E402
    classify_layout_role,
    dump_json,
    ensure_pptx,
    enum_name,
    hex_color,
    parse_brand_md,
    placeholder_idx,
    placeholder_type,
    resolve_layout_map,
    theme_from_part_xml,
)
from colorways import (  # noqa: E402
    apply_colorway,
    designed_theme,
    detect_colorways,
    discover_prototypes,
    extract_named_palette,
)


def _theme_xml(prs: Presentation) -> bytes:
    master = prs.slide_masters[0]
    for rel in master.part.rels.values():
        if rel.reltype == RT.THEME:
            return rel.target_part.blob
    # Fallback: zip theme1.xml
    raise RuntimeError("Theme part not found on slide master.")


def _placeholder_record(shape) -> dict[str, Any]:
    ph_type = placeholder_type(shape)
    return {
        "idx": placeholder_idx(shape),
        "type": int(ph_type) if ph_type is not None else None,
        "type_name": enum_name(ph_type) or None,
        "name": shape.name,
        "left": int(getattr(shape, "left", 0) or 0),
        "top": int(getattr(shape, "top", 0) or 0),
        "width": int(getattr(shape, "width", 0) or 0),
        "height": int(getattr(shape, "height", 0) or 0),
    }


def _run_font(shape) -> dict[str, Any] | None:
    if not getattr(shape, "has_text_frame", False):
        return None
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            size = run.font.size
            color = None
            try:
                color = str(run.font.color.rgb) if run.font.color.rgb else None
            except Exception:
                color = None
            return {
                "size_pt": round(size.pt, 1) if size is not None else None,
                "bold": bool(run.font.bold) if run.font.bold is not None else None,
                "color": hex_color(color, "") if color else None,
                "name": run.font.name,
                "text": (run.text or "")[:80],
            }
    return None


def inspect_template(
    template_path: Path,
    out_dir: Path,
    brand_md: Path | None = None,
    colorway: str | None = None,
) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx_path = ensure_pptx(template_path, out_dir / "TEMPLATE.pptx") if template_path.suffix.lower() == ".potx" else template_path
    if template_path.suffix.lower() == ".potx":
        used = pptx_path
        potx_rewritten = True
    else:
        used = template_path
        potx_rewritten = False
    prs = Presentation(str(used))
    theme = theme_from_part_xml(_theme_xml(prs))
    layouts = []
    for index, layout in enumerate(prs.slide_layouts):
        placeholders = [_placeholder_record(shape) for shape in layout.placeholders]
        layouts.append(
            {
                "index": index,
                "name": layout.name,
                "placeholders": placeholders,
                "role": classify_layout_role(layout.name, placeholders),
            }
        )
    masters = []
    chrome = []
    for mi, master in enumerate(prs.slide_masters):
        master_ph = [_placeholder_record(shape) for shape in master.placeholders]
        master_shapes = []
        for shape in master.shapes:
            is_ph = bool(getattr(shape, "is_placeholder", False))
            rec = {
                "name": shape.name,
                "is_placeholder": is_ph,
                "left": int(getattr(shape, "left", 0) or 0),
                "top": int(getattr(shape, "top", 0) or 0),
                "width": int(getattr(shape, "width", 0) or 0),
                "height": int(getattr(shape, "height", 0) or 0),
                "has_text": bool(getattr(shape, "has_text_frame", False)),
                "has_picture": shape.shape_type is not None and "PICTURE" in str(shape.shape_type),
            }
            master_shapes.append(rec)
            if not is_ph:
                chrome.append(rec)
        masters.append({"index": mi, "placeholders": master_ph, "shape_count": len(master.shapes)})

    sample_title = {"title": None, "subtitle": None}
    layout_title_fonts: dict[str, Any] = {}
    curly = False
    for slide in prs.slides:
        layout = slide.slide_layout
        layout_index = list(prs.slide_layouts).index(layout) if layout in prs.slide_layouts else None
        for shape in slide.placeholders:
            font = _run_font(shape)
            if not font:
                continue
            if font.get("text") and ("“" in font["text"] or "”" in font["text"]):
                curly = True
            ph_type = enum_name(placeholder_type(shape))
            if ph_type in {"CENTER_TITLE", "TITLE"} and sample_title["title"] is None:
                sample_title["title"] = font
            if ph_type == "SUBTITLE" and sample_title["subtitle"] is None:
                sample_title["subtitle"] = font
            if ph_type in {"TITLE", "CENTER_TITLE"} and layout_index is not None:
                layout_title_fonts.setdefault(str(layout_index), font)

    width = int(prs.slide_width)
    height = int(prs.slide_height)
    aspect = "16:9" if abs((width / height) - (16 / 9)) < 0.08 else ("4:3" if abs((width / height) - (4 / 3)) < 0.08 else round(width / height, 4))
    prototypes = discover_prototypes(prs)
    logo_hashes = sorted({h for p in prototypes for h in (p.get("logo_hashes") or []) if h})
    sample_photo_hashes = sorted({h for p in prototypes for h in (p.get("sample_photo_hashes") or []) if h})
    inset_lefts = [int(p["content_box"]["left"]) for p in prototypes if p.get("content_box")]
    inset_tops = [int(p["content_box"]["top"]) for p in prototypes if p.get("content_box")]
    content_inset = {}
    if inset_lefts:
        # Photo layouts park copy in the right half. The page margin is the
        # leftmost text column, not the median of those right-hand wells.
        page_lefts = [left for left in inset_lefts if left < width * 0.45] or list(inset_lefts)
        page_tops = [top for top in inset_tops if top < height * 0.55] or list(inset_tops)
        page_lefts.sort()
        page_tops.sort()
        margin = page_lefts[0]
        content_inset = {
            "left": margin,
            "top": page_tops[len(page_tops) // 2] if page_tops else int(height * 0.16),
            "right": margin,
            "bottom": int(height * 0.08),
        }
    named_layouts = [
        item
        for item in layouts
        if (item.get("name") or "").upper() not in {"DEFAULT", ""}
        and any(
            (ph.get("type_name") or "") in {"BODY", "OBJECT", "TITLE", "CENTER_TITLE"}
            for ph in item.get("placeholders") or []
        )
    ]
    designed_samples = 0
    for slide in prs.slides:
        layout_name = (slide.slide_layout.name or "").strip().upper()
        content_placeholder = False
        for shape in slide.shapes:
            if not getattr(shape, "is_placeholder", False):
                continue
            ph_name = (enum_name(placeholder_type(shape)) or "").upper()
            if ph_name not in {"", "DATE", "FOOTER", "SLIDE_NUMBER"}:
                content_placeholder = True
                break
        if not content_placeholder or layout_name in {"BLANK", "DEFAULT"}:
            designed_samples += 1
    sample_slides_are_the_design = bool(prs.slides) and designed_samples >= int(len(list(prs.slides)) * 0.75)
    if prototypes and (len(layouts) <= 1 or not named_layouts or sample_slides_are_the_design):
        from colorways import layout_map_from_prototypes

        layout_map = layout_map_from_prototypes(prototypes)
        build_mode = "prototypes"
    else:
        layout_map = resolve_layout_map(layouts)
        build_mode = "layouts"
    brand = parse_brand_md(brand_md)
    if brand.get("east_asian_font") and not theme["fonts"].get("minor_ea"):
        theme["fonts"]["minor_ea"] = brand["east_asian_font"]
    colorways = detect_colorways(prs, layouts, theme, prototypes)
    palette = extract_named_palette(prs)
    as_designed = designed_theme(prs, theme)

    tokens = {
        "template_path": str(used.resolve()),
        "source_template_path": str(template_path.resolve()),
        "potx_rewritten": potx_rewritten,
        "slide_width": width,
        "slide_height": height,
        "aspect": aspect,
        "masters": masters,
        "master_chrome": chrome,
        "layouts": layouts,
        "prototypes": prototypes,
        "logo_hashes": logo_hashes,
        "sample_photo_hashes": sample_photo_hashes,
        "content_inset": content_inset,
        "layout_map": colorways[0]["layout_map"] if colorways else layout_map,
        "theme": colorways[0]["theme"] if colorways else as_designed,
        "xml_theme": theme,
        "palette": palette,
        "sample_title_slide_fonts": sample_title,
        "layout_title_fonts": layout_title_fonts,
        "curly_quotes": curly,
        "brand_md": brand,
        "colorways": colorways,
        "default_colorway": (colorways[0]["id"] if colorways else "green"),
        "build_mode": build_mode,
        "notes": [],
    }
    if prototypes:
        if colorways and (colorways[0].get("source") == "accent_pairs"):
            tokens["notes"].append(
                f"Template uses {len(prototypes)} designed sample slides. Convert clones that series."
            )
        elif colorways:
            tokens["notes"].append(
                f"Official template uses {len(prototypes)} sample-slide layouts. "
                f"Dummy layout count={len(layouts)}."
            )
        else:
            swatches = ", ".join(f"{p['label']} {p['hex']}" for p in palette) or "Sawgrass Lime + Water Teal"
            tokens["notes"].append(
                f"Official template uses {len(prototypes)} sample-slide layouts "
                f"({swatches}). Dummy layout count={len(layouts)}."
            )
    if not any(
        ph.get("type_name") in {"BODY", "OBJECT", "VERTICAL_BODY"}
        for layout in layouts
        for ph in layout["placeholders"]
    ) and not prototypes:
        tokens["notes"].append("WARNING: template has zero layouts with a body placeholder.")
    tokens = apply_colorway(tokens, colorway)
    dump_json(out_dir / "brand_tokens.json", tokens)
    return tokens


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect a PPTX/POTX template into brand_tokens.json")
    parser.add_argument("template")
    parser.add_argument("--out-dir", default=".")
    parser.add_argument("--brand-md", default=None)
    parser.add_argument("--colorway", default=None)
    args = parser.parse_args(argv)
    template = Path(args.template)
    if not template.exists():
        print(f"Template not found: {template}", file=sys.stderr)
        return 1
    brand_md = Path(args.brand_md) if args.brand_md else None
    tokens = inspect_template(template, Path(args.out_dir), brand_md, args.colorway)
    print(
        f"Wrote {Path(args.out_dir) / 'brand_tokens.json'} "
        f"({len(tokens['layouts'])} layouts, colorway={tokens.get('colorway') or 'as-designed'})"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
