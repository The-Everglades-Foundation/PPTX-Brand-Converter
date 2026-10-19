"""Green / Blue staff picks mapped onto the 2023 template palette.

The official PPTX contains both Sawgrass Lime (`C1D451`) and Water Teal
(`00ACBF`) as hardcoded sRGB on the sample slides (plus Mangrove Green and
Shell Sand). Those are the source of truth — not Design-tab variants.

Staff still pick one for convert. Green applies Lime across the deck; Blue
applies Teal. Both options stay available. If a later template encodes real
schemes/masters/series, those labels win.
"""

from __future__ import annotations

import io
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Pt

from helpers import (
    A_NS,
    P_NS,
    R_NS,
    apply_picture_crop,
    blob_hash,
    classify_layout_role,
    clamp_box_to_slide,
    contain_fit,
    contain_leftover_fraction,
    content_safe_box,
    cover_crop,
    is_banner_photo_well,
    delete_shape,
    enum_name,
    estimate_lines,
    fits,
    reduce_title_size,
    hex_color,
    is_brand_lockup,
    is_decorative_icon,
    is_org_chrome,
    iter_shapes_abs,
    maybe_recompress,
    overlap_fraction,
    pair_metric_label_items,
    placeholder_type,
    resolve_layout_map,
    rgb,
    theme_color,
    theme_from_part_xml,
)

OFFICE_ACCENT1 = "4472C4"
PALETTE_LABEL_RE = re.compile(
    r"\b([A-Za-z][A-Za-z ]{1,40}?)\s+([0-9A-Fa-f]{6})\b"
)

LIME = "C1D451"
TEAL = "00ACBF"
MANGROVE = "82C769"
SAND = "F1F3E4"
NAVY = "0E0340"
MUTED = "5C5675"
DEEP_TEAL = "00808F"
CREAM = "FDFDFD"
SHELL_LINE = "DFE2CE"

EF_WHITELIST = [
    LIME, TEAL, MANGROVE, SAND, NAVY, MUTED, DEEP_TEAL, CREAM, SHELL_LINE,
    "FFFFFF", "000000", "3B3159", "D6F2F5",
]

SKIP_PROTOTYPE_HINTS = ("notes for staff", "how to use this template", "internal reference")

# Layout labels that belong on the rebuilt slide (not dummy wells).
LAYOUT_CHROME_LABELS = {
    "PRESENTATION TITLE",
    "CONTENTS",
    "THE POINT",
    "BY THE NUMBERS",
    "KICKER",
    "THREE PARTS",
    "TWO POINTS",
    "COMPARISON",
    "SEQUENCE",
    "PEOPLE",
    "NOTES FOR STAFF",
    "BRAND PALETTE",
    "CAPTION",
}
HEADING_DUMMY_LABELS = {
    "bullet one",
    "bullet two",
    "concept one",
    "concept two",
    "concept three",
    "section title here",
    "step title",
    "full name",
    "title, team",
    "row label",
    "label here",
    "column one",
    "column two",
    "column three",
    "column four",
    "phase one",
    "phase two",
    "phase three",
    "phase four",
    "name of speaker",
}
PROTOTYPE_BODY_HINTS = (
    "description text here",
    "two lines of supporting copy",
    "one sentence on what happens here",
    "one line of description for this section",
    "one line explaining what the figure counts",
    "a sentence framing what this section",
    "two or three sentences of body copy",
    "a second short paragraph",
    "section title goes here",
    "a subtitle line goes here",
    "title template",
)
EMU_INCH = 914400


def _base_theme(major: str, minor: str, accent1: str, accent2: str) -> dict[str, Any]:
    return {
        "colors": {
            "dk1": NAVY,
            "lt1": CREAM,
            "dk2": MUTED,
            "lt2": SAND,
            "accent1": accent1,
            "accent2": accent2,
            "accent3": MANGROVE,
            "accent4": DEEP_TEAL,
            "accent5": TEAL,
            "accent6": LIME,
            "hlink": TEAL,
            "folHlink": NAVY,
        },
        "fonts": {
            "major": major or "Questrial",
            "minor": minor or "Arial",
            "major_ea": "",
            "minor_ea": "",
            "major_cs": "",
            "minor_cs": "",
        },
    }


def _sample_fonts(prs) -> tuple[str, str]:
    major = minor = ""
    for slide in list(prs.slides)[:4]:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    name = run.font.name or ""
                    if not name:
                        continue
                    if name.lower() in {"questrial", "poppins"} and not major:
                        major = name
                    if name.lower() in {"arial", "arial nova"} and not minor:
                        minor = name
    return major or "Questrial", minor or "Arial"


def _prototype_role(index: int, title: str, notes: str, total: int) -> str | None:
    blob = f"{title} {notes}".lower()
    if any(hint in blob for hint in SKIP_PROTOTYPE_HINTS):
        return None
    if "cover" in notes.lower() or index == 0:
        return "TITLE"
    if "section divider" in notes.lower() or (title.strip().isdigit() and "section" in blob):
        return "SECTION"
    if "closing" in notes.lower() or "thank you" in title.lower():
        return "CLOSING"
    if "full-bleed" in notes.lower() or "gallery" in notes.lower() or "photograph" in notes.lower() and "caption" in blob:
        return "PICTURE"
    if "two-point" in notes.lower() or "two points" in blob or "three parallel" in notes.lower() or "three parts" in blob:
        return "TWO_CONTENT"
    if "contents" in title.lower() or "standard content" in notes.lower() or "timeline" in notes.lower() or "people" in notes.lower():
        return "TITLE_BODY"
    if "statement" in notes.lower() or "figures" in notes.lower() or "data table" in notes.lower() or "quote" in notes.lower():
        return "TITLE_ONLY"
    if index == total - 1:
        return "CLOSING"
    return classify_layout_role(title, [])


def _shape_box(shape) -> dict[str, int]:
    return {
        "left": int(getattr(shape, "left", 0) or 0),
        "top": int(getattr(shape, "top", 0) or 0),
        "width": int(getattr(shape, "width", 0) or 0),
        "height": int(getattr(shape, "height", 0) or 0),
    }


def _picture_hash(shape) -> str | None:
    try:
        return blob_hash(shape.image.blob)
    except Exception:
        return None


def prototype_picture_records(slide, slide_w: int, slide_h: int) -> list[dict[str, Any]]:
    recs: list[dict[str, Any]] = []
    slide_area = max(1, int(slide_w) * int(slide_h))
    for shape in slide.shapes:
        if shape.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            continue
        box = _shape_box(shape)
        area = max(0, box["width"] * box["height"])
        recs.append(
            {
                **box,
                "area": area,
                "name": shape.name,
                "hash": _picture_hash(shape),
                "logo": is_brand_lockup(
                    area, slide_area, box["left"], box["top"], box["width"], box["height"], slide_w, slide_h
                ),
                "placeholder": _is_picture_placeholder(shape),
            }
        )
    return recs


def prototype_content_box(slide, slide_w: int, slide_h: int, pictures: list[dict[str, Any]]) -> dict[str, int]:
    """Text-column inset for a sample slide, clamped to the slide.

    Full-height side photographs are photo wells, not the text margin.
    """
    boxes: list[dict[str, int]] = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        boxes.append(_shape_box(shape))
    for rec in pictures:
        if rec.get("logo"):
            continue
        if rec["height"] >= slide_h * 0.7 and rec["width"] >= slide_w * 0.3:
            continue
        boxes.append({k: rec[k] for k in ("left", "top", "width", "height")})
    if not boxes:
        side = int(slide_w * 0.05)
        top = int(slide_h * 0.16)
        return {"left": side, "top": top, "width": slide_w - 2 * side, "height": slide_h - top - int(slide_h * 0.08)}
    left = min(int(b["left"]) for b in boxes)
    top = min(int(b["top"]) for b in boxes)
    right = max(int(b["left"]) + int(b["width"]) for b in boxes)
    bottom = max(int(b["top"]) + int(b["height"]) for b in boxes)
    return clamp_box_to_slide({"left": left, "top": top, "width": right - left, "height": bottom - top}, slide_w, slide_h)


def discover_prototypes(prs) -> list[dict[str, Any]]:
    prototypes = []
    total = len(prs.slides)
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    for index, slide in enumerate(prs.slides):
        notes = ""
        if slide.has_notes_slide:
            notes = (slide.notes_slide.notes_text_frame.text or "").strip()
        title = ""
        texts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                t = (shape.text_frame.text or "").strip()
                if t:
                    texts.append(t.split("\n")[0])
        title = texts[1] if len(texts) > 1 else (texts[0] if texts else f"Slide {index + 1}")
        if texts and texts[0].isupper() and len(texts) > 1:
            name = texts[0].title()
        else:
            name = title[:60] or f"Slide {index + 1}"
        role = _prototype_role(index, " ".join(texts[:3]), notes, total)
        if role is None:
            continue
        pictures = prototype_picture_records(slide, slide_w, slide_h)
        frames = [p for p in pictures if not p.get("logo")]
        prototypes.append(
            {
                "index": index,
                "name": name,
                "role": role,
                "notes": notes.split("\n")[0][:160],
                "photo_frames": [
                    {k: p[k] for k in ("left", "top", "width", "height", "hash") if k in p}
                    for p in frames
                ],
                "photo_frame_count": len(frames),
                "logo_hashes": [p["hash"] for p in pictures if p.get("logo") and p.get("hash")],
                "sample_photo_hashes": [p["hash"] for p in frames if p.get("hash")],
                "content_box": prototype_content_box(slide, slide_w, slide_h, pictures),
            }
        )
    return prototypes


def layout_map_from_prototypes(prototypes: list[dict[str, Any]]) -> dict[str, int]:
    fake_layouts = [{"index": p["index"], "name": p["name"], "role": p["role"], "placeholders": [{"type_name": "TITLE"}]} for p in prototypes]
    # BLANK: use a content-like prototype if no blank exists
    mapping = resolve_layout_map(fake_layouts)
    return mapping


def _is_generic_office_theme(theme: dict[str, Any] | None) -> bool:
    colors = (theme or {}).get("colors") or {}
    fonts = (theme or {}).get("fonts") or {}
    accent1 = hex_color(colors.get("accent1") or "")
    major = (fonts.get("major") or "").lower()
    return accent1 == OFFICE_ACCENT1 or major.startswith("calibri")


def designed_theme(prs, xml_theme: dict[str, Any] | None = None) -> dict[str, Any]:
    """Theme tokens taken from the template's sample slides + staff palette.

    The package theme part is the leftover Office theme (Calibri / 4472C4). Brand
    color lives as hardcoded sRGB on the prototypes, so inspect uses that palette
    for charts/tables instead of the unused Office scheme.
    """
    major, minor = _sample_fonts(prs)
    if xml_theme and not _is_generic_office_theme(xml_theme):
        themed = deepcopy(xml_theme)
        fonts = dict(themed.get("fonts") or {})
        fonts["major"] = major or fonts.get("major") or "Questrial"
        fonts["minor"] = minor or fonts.get("minor") or "Arial"
        themed["fonts"] = fonts
        return themed
    return _base_theme(major, minor, LIME, TEAL)


def extract_named_palette(prs) -> list[dict[str, str]]:
    """Named swatches from the staff 'BRAND PALETTE' sample slide, if present."""
    found: list[dict[str, str]] = []
    seen: set[str] = set()
    for slide in prs.slides:
        texts: list[str] = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                texts.append(shape.text_frame.text or "")
        blob = "\n".join(texts)
        if "brand palette" not in blob.lower() and "sawgrass" not in blob.lower():
            continue
        for match in PALETTE_LABEL_RE.finditer(blob):
            label = re.sub(r"\s+", " ", match.group(1)).strip()
            hex_value = hex_color(match.group(2))
            if hex_value in seen:
                continue
            seen.add(hex_value)
            found.append({"label": label, "hex": hex_value})
    return found


def _layout_names_for(layouts: list[dict[str, Any]], prototypes: list[dict[str, Any]], layout_map: dict[str, int]) -> dict[str, str]:
    names: dict[str, str] = {}
    for proto in prototypes:
        names[str(proto["index"])] = proto["name"]
    for layout in layouts:
        names.setdefault(str(layout["index"]), layout.get("name"))
    return {role: names.get(str(idx), role) for role, idx in layout_map.items()}


def _shared_maps(layouts: list[dict[str, Any]], prototypes: list[dict[str, Any]]) -> dict[str, Any]:
    proto_map = layout_map_from_prototypes(prototypes) if prototypes else resolve_layout_map(layouts)
    return {
        "layout_map": proto_map,
        "layout_names": _layout_names_for(layouts, prototypes, proto_map),
        "prototype_indices": [p["index"] for p in prototypes],
        "build_mode": "prototypes" if prototypes else "layouts",
        "srgb_map": {},
    }


def _scheme_from_element(scheme) -> dict[str, str]:
    colors: dict[str, str] = {}
    if scheme is None:
        return colors
    for child in list(scheme):
        name = child.tag.split("}")[-1]
        srgb = child.find(f".//{{{A_NS}}}srgbClr")
        sys_clr = child.find(f".//{{{A_NS}}}sysClr")
        if srgb is not None and srgb.get("val"):
            colors[name] = hex_color(srgb.get("val"))
        elif sys_clr is not None:
            colors[name] = hex_color(sys_clr.get("lastClr") or "")
    return {k: v for k, v in colors.items() if v}


def _slug_id(label: str, fallback: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", (label or "").lower()).strip("-")
    return slug or fallback


def _colorways_from_extra_schemes(prs, layouts, xml_theme, prototypes) -> list[dict[str, Any]]:
    options: list[dict[str, Any]] = []
    seen: set[str] = set()
    shared = _shared_maps(layouts, prototypes)
    base = designed_theme(prs, xml_theme)
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    except Exception:
        return options
    theme_blobs: list[bytes] = []
    for master in prs.slide_masters:
        for rel in master.part.rels.values():
            if rel.reltype == RT.THEME:
                theme_blobs.append(rel.target_part.blob)
    for blob in theme_blobs:
        try:
            root = etree.fromstring(blob)
        except Exception:
            continue
        extra = root.find(f".//{{{A_NS}}}extraClrSchemeLst")
        if extra is None:
            continue
        for extra_scheme in extra:
            scheme = extra_scheme.find(f"{{{A_NS}}}clrScheme")
            if scheme is None:
                continue
            name = (scheme.get("name") or "").strip()
            colors = _scheme_from_element(scheme)
            if not name or not colors:
                continue
            option_id = _slug_id(name, f"scheme-{len(options)+1}")
            if option_id in seen:
                continue
            seen.add(option_id)
            themed = deepcopy(base)
            themed["colors"] = {**(themed.get("colors") or {}), **colors}
            options.append(
                {
                    "id": option_id,
                    "label": name,
                    "description": f"Theme color scheme “{name}” from the template",
                    "source": "extra_clr_scheme",
                    "theme": themed,
                    **shared,
                }
            )
    return options


def _colorways_from_masters(prs, layouts, xml_theme, prototypes) -> list[dict[str, Any]]:
    masters = list(prs.slide_masters)
    if len(masters) < 2:
        return []
    options: list[dict[str, Any]] = []
    base = designed_theme(prs, xml_theme)
    for index, master in enumerate(masters):
        name = (getattr(master, "name", None) or "").strip() or f"Master {index + 1}"
        master_layouts = []
        try:
            for li, layout in enumerate(master.slide_layouts):
                placeholders = []
                master_layouts.append(
                    {
                        "index": li,
                        "name": layout.name,
                        "placeholders": placeholders,
                        "role": classify_layout_role(layout.name, []),
                    }
                )
        except Exception:
            master_layouts = [item for item in layouts if item.get("master_index") == index] or layouts
        layout_map = resolve_layout_map(master_layouts) if master_layouts else resolve_layout_map(layouts)
        options.append(
            {
                "id": _slug_id(name, f"master-{index+1}"),
                "label": name,
                "description": f"Slide master “{name}”",
                "source": "slide_master",
                "master_index": index,
                "theme": deepcopy(base),
                "layout_map": layout_map,
                "layout_names": {role: next((l["name"] for l in master_layouts if l["index"] == idx), role) for role, idx in layout_map.items()},
                "prototype_indices": [p["index"] for p in prototypes],
                "build_mode": "prototypes" if prototypes else "layouts",
                "srgb_map": {},
            }
        )
    return options


def _colorways_from_prototype_series(prs, layouts, xml_theme, prototypes) -> list[dict[str, Any]]:
    """Split sample slides only when notes/titles name distinct green vs blue series."""
    buckets: dict[str, list[dict[str, Any]]] = {"green": [], "blue": []}
    for proto in prototypes:
        blob = f"{proto.get('name') or ''} {proto.get('notes') or ''}".lower()
        if re.search(r"\b(blue|teal|water teal)\b", blob) and not re.search(r"\b(green|lime|sawgrass)\b", blob):
            buckets["blue"].append(proto)
        elif re.search(r"\b(green|lime|sawgrass)\b", blob) and not re.search(r"\b(blue|teal)\b", blob):
            buckets["green"].append(proto)
    if not buckets["green"] or not buckets["blue"]:
        return []
    # Require overlapping roles so they are variants of the same layouts, not mixed accents.
    green_roles = {p["role"] for p in buckets["green"] if p.get("role")}
    blue_roles = {p["role"] for p in buckets["blue"] if p.get("role")}
    if not (green_roles & blue_roles):
        return []
    base = designed_theme(prs, xml_theme)
    options = []
    labels = {"green": "Green", "blue": "Blue"}
    for key, group in buckets.items():
        layout_map = layout_map_from_prototypes(group)
        options.append(
            {
                "id": key,
                "label": labels[key],
                "description": f"Sample-slide series labeled {labels[key]} in the template",
                "source": "prototype_series",
                "theme": deepcopy(base),
                "layout_map": layout_map,
                "layout_names": _layout_names_for(layouts, group, layout_map),
                "prototype_indices": [p["index"] for p in group],
                "build_mode": "prototypes",
                "srgb_map": {},
            }
        )
    return options


def _colorways_from_theme_overrides(prs, layouts, xml_theme, prototypes) -> list[dict[str, Any]]:
    options: list[dict[str, Any]] = []
    shared = _shared_maps(layouts, prototypes)
    base = designed_theme(prs, xml_theme)
    pkg = getattr(getattr(prs, "part", None), "package", None)
    if pkg is None:
        return options
    try:
        parts = list(pkg.iter_parts())
    except Exception:
        return options
    seen: set[str] = set()
    for part in parts:
        name = str(getattr(part, "partname", "") or "")
        if "themeOverride" not in name and "themeoverride" not in name.lower():
            continue
        try:
            blob = part.blob
            parsed = theme_from_part_xml(blob)
        except Exception:
            continue
        colors = parsed.get("colors") or {}
        if not colors:
            continue
        label = Path(name).stem
        option_id = _slug_id(label, f"override-{len(options)+1}")
        if option_id in seen:
            continue
        seen.add(option_id)
        themed = deepcopy(base)
        themed["colors"] = {**(themed.get("colors") or {}), **colors}
        options.append(
            {
                "id": option_id,
                "label": label,
                "description": f"Theme override {label}",
                "source": "theme_override",
                "theme": themed,
                **shared,
            }
        )
    return options


def _palette_swatch(palette: list[dict[str, str]], hex_value: str, fallback: str) -> str:
    for item in palette:
        if hex_color(item.get("hex") or "") == hex_color(hex_value):
            return item.get("label") or fallback
    return fallback


def _colorways_from_palette(prs, layouts, xml_theme, prototypes) -> list[dict[str, Any]]:
    """Green / Blue from the template's own Sawgrass Lime and Water Teal swatches."""
    palette = extract_named_palette(prs)
    lime_label = _palette_swatch(palette, LIME, "Sawgrass Lime")
    teal_label = _palette_swatch(palette, TEAL, "Water Teal")
    shared = _shared_maps(layouts, prototypes)
    shared.pop("srgb_map", None)
    major, minor = _sample_fonts(prs)
    green_theme = _base_theme(major, minor, LIME, TEAL)
    blue_theme = _base_theme(major, minor, TEAL, LIME)
    return [
        {
            "id": "green",
            "label": "Green",
            "description": f"{lime_label} {LIME}",
            "source": "palette",
            "theme": green_theme,
            "srgb_map": {TEAL: LIME, DEEP_TEAL: MANGROVE},
            **shared,
        },
        {
            "id": "blue",
            "label": "Blue",
            "description": f"{teal_label} {TEAL}",
            "source": "palette",
            "theme": blue_theme,
            "srgb_map": {LIME: TEAL, MANGROVE: DEEP_TEAL},
            **shared,
        },
    ]


def detect_colorways(
    prs,
    layouts: list[dict[str, Any]],
    theme: dict[str, Any],
    prototypes: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Return Green/Blue (or later structural options) for the staff picker.

    Always includes both template palette colors as pickable colorways.
    """
    for finder in (
        _colorways_from_extra_schemes,
        _colorways_from_masters,
        _colorways_from_theme_overrides,
        _colorways_from_prototype_series,
    ):
        found = finder(prs, layouts, theme, prototypes)
        if len(found) >= 2:
            return found
    return _colorways_from_palette(prs, layouts, theme, prototypes)


def apply_colorway(tokens: dict[str, Any], colorway_id: str | None) -> dict[str, Any]:
    tokens = dict(tokens)
    colorways = tokens.get("colorways") or []
    brand = dict(tokens.get("brand_md") or {})
    brand["color_whitelist"] = sorted(set((brand.get("color_whitelist") or []) + EF_WHITELIST))
    extra_fonts = [
        ((tokens.get("theme") or {}).get("fonts") or {}).get("major"),
        ((tokens.get("theme") or {}).get("fonts") or {}).get("minor"),
        "Questrial",
        "Arial",
        "Arial Nova",
        "Poppins",
    ]
    brand["fonts"] = sorted({f for f in list(brand.get("fonts") or []) + extra_fonts if f})
    tokens["brand_md"] = brand
    tokens.setdefault("srgb_map", {})

    if not colorways:
        tokens["colorway"] = "green"
        tokens["colorway_label"] = "Green"
        tokens["srgb_map"] = {}
        return tokens

    requested = (colorway_id or tokens.get("colorway") or tokens.get("default_colorway") or "green").strip().lower()
    match = next((c for c in colorways if str(c.get("id") or "").lower() == requested), None)
    if match is None:
        aliases = {"lime": "green", "sawgrass": "green", "sawgrass-lime": "green", "teal": "blue", "water-teal": "blue"}
        alias = aliases.get(requested)
        if alias:
            match = next((c for c in colorways if str(c.get("id") or "").lower() == alias), None)
    if match is None:
        match = colorways[0]
    tokens["colorway"] = match["id"]
    tokens["theme"] = match.get("theme") or tokens.get("theme")
    tokens["layout_map"] = match.get("layout_map") or tokens.get("layout_map")
    tokens["layout_names"] = match.get("layout_names") or tokens.get("layout_names")
    tokens["srgb_map"] = match.get("srgb_map") or {}
    tokens["build_mode"] = match.get("build_mode") or tokens.get("build_mode") or "layouts"
    tokens["colorway_label"] = match.get("label") or match["id"]
    extra_fonts.append(((match.get("theme") or {}).get("fonts") or {}).get("major"))
    extra_fonts.append(((match.get("theme") or {}).get("fonts") or {}).get("minor"))
    brand["fonts"] = sorted({f for f in list(brand.get("fonts") or []) + extra_fonts if f})
    tokens["brand_md"] = brand
    return tokens


def remap_srgb_in_slide(slide, srgb_map: dict[str, str]) -> None:
    if not srgb_map:
        return
    mapping = {hex_color(k): hex_color(v) for k, v in srgb_map.items()}
    for node in slide._element.findall(f".//{{{A_NS}}}srgbClr"):
        val = hex_color(node.get("val") or "")
        if val in mapping:
            node.set("val", mapping[val])


def patch_theme_part(prs, theme: dict[str, Any]) -> None:
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    colors = (theme or {}).get("colors") or {}
    fonts = (theme or {}).get("fonts") or {}
    master = prs.slide_masters[0]
    part = None
    for rel in master.part.rels.values():
        if rel.reltype == RT.THEME:
            part = rel.target_part
            break
    if part is None:
        return
    root = etree.fromstring(part.blob)
    scheme = root.find(f".//{{{A_NS}}}clrScheme")
    if scheme is not None:
        for name, hex_value in colors.items():
            node = scheme.find(f"{{{A_NS}}}{name}")
            if node is None:
                continue
            for child in list(node):
                node.remove(child)
            etree.SubElement(node, f"{{{A_NS}}}srgbClr").set("val", hex_color(hex_value))
    if fonts.get("major") or fonts.get("minor"):
        major = root.find(f".//{{{A_NS}}}majorFont")
        minor = root.find(f".//{{{A_NS}}}minorFont")
        if major is not None and fonts.get("major"):
            latin = major.find(f"{{{A_NS}}}latin")
            if latin is not None:
                latin.set("typeface", fonts["major"])
        if minor is not None and fonts.get("minor"):
            latin = minor.find(f"{{{A_NS}}}latin")
            if latin is not None:
                latin.set("typeface", fonts["minor"])
    part.blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def duplicate_slide(prs, index: int):
    """Clone a template sample slide inside the same package (not from SOURCE)."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)
    src_cSld = source._element.find(qn("p:cSld"))
    dest_cSld = dest._element.find(qn("p:cSld"))
    if src_cSld is not None and dest_cSld is not None:
        src_bg = src_cSld.find(qn("p:bg"))
        dest_bg = dest_cSld.find(qn("p:bg"))
        if src_bg is not None:
            if dest_bg is not None:
                dest_cSld.remove(dest_bg)
            copied = deepcopy(src_bg)
            dest_cSld.insert(0, copied)
            _rewire_blips(copied, source.part, dest)
    src_ovr = source._element.find(qn("p:clrMapOvr"))
    dest_ovr = dest._element.find(qn("p:clrMapOvr"))
    if src_ovr is not None:
        copied_ovr = deepcopy(src_ovr)
        if dest_ovr is not None:
            dest._element.replace(dest_ovr, copied_ovr)
        else:
            dest._element.append(copied_ovr)
    for shape in list(dest.shapes):
        el = shape._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    source_part = source.part
    for shape in source.shapes:
        new_el = deepcopy(shape._element)
        _rewire_blips(new_el, source_part, dest)
        dest.shapes._spTree.append(new_el)
    return dest


def _rewire_blips(element, source_part, dest_slide) -> None:
    embed_attr = f"{{{R_NS}}}embed"
    for blip in element.findall(f".//{{{A_NS}}}blip"):
        rid = blip.get(embed_attr)
        if not rid:
            continue
        try:
            image_part = source_part.related_part(rid)
            blob = image_part.blob
        except Exception:
            continue
        try:
            _image_part, new_rid = dest_slide.part.get_or_add_image_part(io.BytesIO(blob))
            blip.set(embed_attr, new_rid)
        except Exception:
            continue


def prototype_text_slots(slide) -> list[dict[str, Any]]:
    slots = []
    for shape, left, top in iter_shapes_abs(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        slots.append(
            {
                "shape": shape,
                "left": left,
                "top": top,
                "width": int(shape.width),
                "height": int(shape.height),
                "text": text,
            }
        )
    slots.sort(key=lambda s: (int(s["top"]), int(s["left"])))
    return slots


def is_short_caption(text: str) -> bool:
    """True for a one- or two-line table caption, not competing body copy."""
    t = (text or "").strip()
    if not t or is_layout_chrome(t):
        return False
    lines = [ln for ln in t.splitlines() if ln.strip()]
    if len(lines) > 2:
        return False
    return 1 <= len(t.split()) <= 40


def is_layout_chrome(text: str) -> bool:
    """True for brand lockup labels / index badges that must stay on the sample slide."""
    t = (text or "").strip()
    if not t:
        return False
    if t in LAYOUT_CHROME_LABELS:
        return True
    if re.fullmatch(r"0?\d{1,2}", t):
        return True
    if t in {'"', "“", "”", "'", "‘", "’"}:
        return True
    if "evergladesfoundation.org" in t.lower():
        return True
    return False


def is_leftover_prototype_copy(text: str) -> bool:
    """True for unused sample-slide dummy that must not survive convert."""
    t = (text or "").strip()
    if not t or is_layout_chrome(t):
        return False
    low = t.lower()
    if low in HEADING_DUMMY_LABELS:
        return True
    return any(hint in low for hint in PROTOTYPE_BODY_HINTS)


def _write_paragraph_text(para, text: str) -> None:
    if para.runs:
        para.runs[0].text = text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.text = text
    nodes = para._p.findall(f".//{{{A_NS}}}t")
    if nodes:
        nodes[0].text = text
        for node in nodes[1:]:
            node.text = ""
    elif text:
        para.text = text


def replace_shape_text(shape, text: str) -> None:
    """Replace visible copy and wipe leftover prototype paragraphs in the same frame."""
    if not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.word_wrap = True
    parts = [p for p in str(text).split("\n")] if text else [""]
    if not parts:
        parts = [""]
    while len(tf.paragraphs) < len(parts):
        tf.add_paragraph()
    for i, para in enumerate(tf.paragraphs):
        value = parts[i] if i < len(parts) else ""
        _write_paragraph_text(para, value)


def _apply_font_size(shape, size_pt: float) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            try:
                run.font.size = Pt(size_pt)
            except Exception:
                continue


def fit_title_in_slot(shape, text: str) -> None:
    """Shrink a prototype title so ALL-CAPS words stay on one line (no orphan letter)."""
    if not getattr(shape, "has_text_frame", False) or not (text or "").strip():
        return
    box_w = int(shape.width or 0)
    box_h = int(shape.height or 0)
    if box_w <= 0 or box_h <= 0:
        return
    start = int(_slot_font_pt({"shape": shape, "height": box_h}) or 40)
    cover_well = box_h >= int(2.0 * EMU_INCH)
    size = reduce_title_size(
        text,
        box_w,
        box_h,
        start,
        max_lines=2 if cover_well else None,
        fit_height=cover_well,
    )
    _apply_font_size(shape, size)
    # One-word covers must never wrap the last letter onto its own line.
    if len((text or "").split()) == 1:
        try:
            shape.text_frame.word_wrap = False
        except Exception:
            pass


def clear_shape_text(shape) -> None:
    replace_shape_text(shape, "")


def _looks_like_heading_slot(slot: dict[str, Any]) -> bool:
    t = (slot.get("text") or "").strip()
    if t.lower() in HEADING_DUMMY_LABELS:
        return True
    height = int(slot.get("height") or 0)
    if height <= int(0.7 * EMU_INCH) and len(t) <= 48 and "\n" not in t:
        return True
    return False


def _cluster_columns(slots: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    if not slots:
        return []
    ordered = sorted(slots, key=lambda s: (int(s["left"]), int(s["top"])))
    groups: list[list[dict[str, Any]]] = [[ordered[0]]]
    for slot in ordered[1:]:
        col = groups[-1]
        col_left = min(int(s["left"]) for s in col)
        col_right = max(int(s["left"]) + int(s["width"]) for s in col)
        center = int(slot["left"]) + int(slot["width"]) / 2
        col_center = (col_left + col_right) / 2
        if int(slot["left"]) >= col_right - int(0.2 * EMU_INCH) and center - col_center >= 1.2 * EMU_INCH:
            groups.append([slot])
        else:
            col.append(slot)
    return groups


_SPLIT_MARKER_RE = re.compile(r"^(?:(?:step|phase)\s+\d+\b|\d+\s*[·.)-])", re.I)


def _is_split_heading(text: str) -> bool:
    compact = " ".join((text or "").split())
    if not compact:
        return False
    if _SPLIT_MARKER_RE.match(compact):
        return True
    if " — " in compact or " – " in compact:
        label = compact.split(" — ")[0].split(" – ")[0]
        return 1 <= len(label.split()) <= 8
    return 1 <= len(compact.split()) <= 8 and not compact.endswith((".", "!", "?"))


def _heading_group_starts(items: list[str]) -> list[int]:
    markers = [i for i, text in enumerate(items) if _SPLIT_MARKER_RE.match(" ".join(text.split()))]
    starts: list[int] = []
    if len(markers) >= 2:
        starts.extend(markers)
        for i, text in enumerate(items):
            if i in starts or (i and (i - 1) in starts):
                continue
            compact = " ".join(text.split())
            if " — " in compact or " – " in compact:
                continue
            if 1 <= len(compact.split()) <= 6 and not compact.endswith((".", "!", "?")):
                starts.append(i)
        return sorted(starts)
    for i, text in enumerate(items):
        if _is_split_heading(text):
            starts.append(i)
    return starts


def _distribute_groups(groups: list[list[str]], n: int) -> list[list[str]]:
    if n <= 1:
        return [[part for group in groups for part in group]]
    if len(groups) <= n:
        return groups + [[] for _ in range(n - len(groups))]
    head = groups[: n - 1]
    tail = [part for group in groups[n - 1 :] for part in group]
    return head + [tail]


def _split_into(items: list[str], n: int) -> list[list[str]]:
    if n <= 1:
        return [items]
    metric_groups = pair_metric_label_items(items)
    if metric_groups:
        return _distribute_groups(metric_groups, n)
    starts = _heading_group_starts(items)
    if len(starts) >= 2:
        groups = []
        for i, start in enumerate(starts):
            end = starts[i + 1] if i + 1 < len(starts) else len(items)
            groups.append(items[start:end])
        if starts[0] > 0:
            groups[0] = items[: starts[0]] + groups[0]
        return _distribute_groups(groups, n)
    if n > 1 and len(items) == n * 3:
        return [items[i * 3 : (i + 1) * 3] for i in range(n)]
    if n > 1 and len(items) == n * 2:
        return [items[i * 2 : (i + 1) * 2] for i in range(n)]
    if len(items) >= n * 2 and len(items) % 2 == 0 and len(items) % 3 != 0:
        pairs = list(zip(items[0::2], items[1::2]))
        base, extra = divmod(len(pairs), n)
        chunks: list[list[str]] = []
        i = 0
        for ci in range(n):
            take = base + (1 if ci < extra else 0)
            chunk_pairs = pairs[i : i + take]
            i += take
            chunks.append([part for pair in chunk_pairs for part in pair])
        return chunks
    base, extra = divmod(len(items), n)
    chunks = []
    i = 0
    for ci in range(n):
        take = base + (1 if ci < extra else 0)
        chunks.append(items[i : i + take])
        i += take
    return chunks


def _fill_body_stack(slots: list[dict[str, Any]], texts: list[str]) -> list[Any]:
    """Write body into one well; delete stacked dummy frames in the same column."""
    to_delete: list[Any] = []
    if not slots:
        return to_delete
    slots = sorted(slots, key=lambda s: (int(s["top"]), int(s["left"])))
    # Prefer the roomiest well so collapsed copy uses body type, not 33pt heading.
    primary = max(slots, key=lambda s: int(s.get("height") or 0) * int(s.get("width") or 0))
    if len(slots) > 1:
        left = min(int(s["left"]) for s in slots)
        top = min(int(s["top"]) for s in slots)
        right = max(int(s["left"]) + int(s["width"]) for s in slots)
        bottom = max(int(s["top"]) + int(s["height"]) for s in slots)
        try:
            primary["shape"].left = int(left)
            primary["shape"].top = int(top)
            primary["shape"].width = int(right - left)
            primary["shape"].height = int(bottom - top)
        except Exception:
            pass
        for extra in slots:
            if extra is primary:
                continue
            clear_shape_text(extra["shape"])
            to_delete.append(extra["shape"])
    if texts:
        replace_shape_text(primary["shape"], "\n".join(texts))
    else:
        clear_shape_text(primary["shape"])
        to_delete.append(primary["shape"])
    return to_delete


def _slot_font_pt(slot: dict[str, Any]) -> float:
    shape = slot.get("shape")
    if shape is not None and getattr(shape, "has_text_frame", False):
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size is not None:
                    return float(run.font.size.pt)
    height = int(slot.get("height") or 0)
    return 33.0 if height <= int(0.7 * EMU_INCH) else 24.0


def _heading_text_fits(text: str, slot: dict[str, Any]) -> bool:
    """Long takeaway sentences must not be forced into a 0.55″ prototype heading."""
    t = (text or "").strip()
    if not t:
        return False
    height = int(slot.get("height") or 0)
    width = int(slot.get("width") or 0)
    if height < int(0.35 * EMU_INCH) or width <= 0:
        return False
    font_pt = _slot_font_pt(slot)
    lines = estimate_lines(t, font_pt, width)
    if lines <= 1:
        # Sample heading wells are one line of 33pt; 1.2 body leading is tighter than the box.
        return font_pt * (EMU_INCH / 72) <= height * 1.2
    return fits([t], font_pt, width, height)


def _peel_lead_in(body_texts: list[str]) -> tuple[str, list[str]]:
    """Full-width intro sentences belong under the title, not in KPI card 1."""
    if len(body_texts) < 5:
        return "", body_texts
    first = body_texts[0]
    if len(first.split()) < 16:
        return "", body_texts
    rest = body_texts[1:]
    shortish = sum(1 for item in rest if len(item.split()) <= 10)
    if shortish >= 4:
        return first, rest
    return "", body_texts


def _fill_column(slots: list[dict[str, Any]], texts: list[str]) -> list[Any]:
    if not slots:
        return []
    slots = sorted(slots, key=lambda s: (int(s["top"]), int(s["left"])))
    if not texts:
        deleted = []
        for slot in slots:
            clear_shape_text(slot["shape"])
            deleted.append(slot["shape"])
        return deleted
    heading = slots[0]
    if (
        _looks_like_heading_slot(heading)
        and len(slots) >= 2
        and _heading_text_fits(texts[0], heading)
    ):
        replace_shape_text(heading["shape"], texts[0])
        return _fill_body_stack(slots[1:], texts[1:])
    # One well per card: long copy + leftover dummy frames collapse together
    # so heading overflow cannot stack on a second text layer.
    return _fill_body_stack(slots, texts)


def _wide_slot(slots: list[dict[str, Any]]) -> dict[str, Any] | None:
    for slot in slots:
        if int(slot.get("width") or 0) >= int(8.0 * EMU_INCH) and int(slot.get("height") or 0) <= int(1.4 * EMU_INCH):
            return slot
    return None


def expand_text_wells(slide) -> None:
    """Grow filled body wells into unused space so long copy does not clip."""
    frames: list[tuple[Any, dict[str, int], str]] = []
    pics: list[dict[str, int]] = []
    slide_h = int(11.25 * EMU_INCH)
    try:
        slide_h = int(slide.part.package.presentation_part.presentation.slide_height)
    except Exception:
        pass
    for shape in slide.shapes:
        box = _shape_box(shape)
        if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            pics.append(box)
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text or is_layout_chrome(text):
            continue
        frames.append((shape, box, text))
    if not frames:
        return
    title_top = min(int(b["top"]) for _s, b, _t in frames)
    ceiling = int(slide_h - 0.35 * EMU_INCH)
    wide_pics = [p for p in pics if int(p["width"]) >= int(4.0 * EMU_INCH)]
    if wide_pics:
        banner = min(int(p["top"]) for p in wide_pics) - int(0.16 * EMU_INCH)
        if banner > title_top + int(1.0 * EMU_INCH):
            ceiling = min(ceiling, banner)
    if ceiling <= 0:
        return
    for shape, box, text in frames:
        if int(box["top"]) <= title_top + int(0.05 * EMU_INCH) and int(box["width"]) >= int(8.0 * EMU_INCH):
            continue
        if len(text.split()) <= 3 and int(box["height"]) <= int(0.7 * EMU_INCH):
            continue
        if is_short_caption(text) and int(box["width"]) >= int(12.0 * EMU_INCH):
            continue
        col_left = int(box["left"]) - int(0.35 * EMU_INCH)
        col_right = int(box["left"]) + int(box["width"]) + int(0.35 * EMU_INCH)
        next_top = ceiling
        for _other, other_box, _ot in frames:
            if _other is shape:
                continue
            other_cx = int(other_box["left"]) + int(other_box["width"]) / 2
            if col_left <= other_cx <= col_right and int(other_box["top"]) > int(box["top"]) + int(0.2 * EMU_INCH):
                next_top = min(next_top, int(other_box["top"]) - int(0.12 * EMU_INCH))
        new_bottom = next_top
        new_height = new_bottom - int(box["top"])
        if new_height > int(box["height"]) + int(0.15 * EMU_INCH):
            try:
                shape.height = int(new_height)
            except Exception:
                continue


def _looks_like_comparison_grid(slots: list[dict[str, Any]]) -> bool:
    """Comparison sample art is a dense row/column dummy grid, not two cards."""
    if len(slots) < 8:
        return False
    heights = [int(s.get("height") or 0) for s in slots]
    median_h = sorted(heights)[len(heights) // 2]
    similar = sum(1 for h in heights if abs(h - median_h) <= int(0.35 * EMU_INCH))
    return similar >= 6


def _fill_caption_on_grid(slots: list[dict[str, Any]], texts: list[str]) -> list[Any]:
    """Place a short caption under the title; do not merge dummy table cells."""
    to_delete: list[Any] = []
    if not slots:
        return to_delete
    slots = sorted(slots, key=lambda s: (int(s["top"]), int(s["left"])))
    primary = slots[0]
    left = min(int(s["left"]) for s in slots)
    top = min(int(s["top"]) for s in slots)
    right = max(int(s["left"]) + int(s["width"]) for s in slots)
    cap_h = int(0.55 * EMU_INCH)
    if texts and len(" ".join(texts).split()) > 18:
        cap_h = int(0.78 * EMU_INCH)
    try:
        primary["shape"].left = int(left)
        primary["shape"].top = int(top)
        primary["shape"].width = max(int(1.0 * EMU_INCH), int(right - left))
        primary["shape"].height = int(cap_h)
    except Exception:
        pass
    if texts:
        replace_shape_text(primary["shape"], "\n".join(texts))
        try:
            primary["shape"].fill.background()
        except Exception:
            pass
    else:
        clear_shape_text(primary["shape"])
        to_delete.append(primary["shape"])
    for extra in slots[1:]:
        clear_shape_text(extra["shape"])
        to_delete.append(extra["shape"])
    return to_delete


def _recolor_text_shape(shape, color_hex: str, size_pt: float | None = None) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    for para in shape.text_frame.paragraphs:
        if not para.runs:
            run = para.add_run()
            run.text = para.text or ""
        for run in para.runs:
            try:
                run.font.color.rgb = rgb(color_hex)
            except Exception:
                pass
            if size_pt:
                try:
                    run.font.size = Pt(size_pt)
                except Exception:
                    pass


def prepare_table_caption_wells(slide, tokens: dict[str, Any]) -> None:
    """Short table captions sit under the title in dark ink, not a white overlay well."""
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    safe = content_safe_box(tokens, tokens.get("_layout_index"))
    titles: list[tuple[Any, dict[str, int]]] = []
    captions: list[tuple[Any, dict[str, int], str]] = []
    other = False
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        box = _shape_box(shape)
        if is_layout_chrome(text):
            titles.append((shape, box))
            continue
        if box["top"] < slide_h * 0.28 and box["width"] >= slide_w * 0.45 and box["height"] <= int(1.4 * EMU_INCH):
            titles.append((shape, box))
            continue
        if is_short_caption(text):
            captions.append((shape, box, text))
        else:
            other = True
    if other or not captions:
        return
    title_bottom = int(safe["top"])
    for _shape, box in titles:
        title_bottom = max(title_bottom, int(box["top"]) + int(box["height"]))
    gap = int(0.12 * EMU_INCH)
    y = title_bottom + gap
    words = max(len(t.split()) for _s, _b, t in captions)
    cap_h = int(0.78 * EMU_INCH) if words > 18 else int(0.52 * EMU_INCH)
    ink = theme_color(tokens, "dk2", "3B3159")
    if hex_color(ink) in {"FFFFFF", "FDFDFD", "F1F3E4"}:
        ink = "3B3159"
    kept = captions[0][0]
    try:
        kept.left = int(safe["left"])
        kept.top = int(y)
        kept.width = int(safe["width"])
        kept.height = int(cap_h)
    except Exception:
        pass
    try:
        kept.fill.background()
    except Exception:
        pass
    _recolor_text_shape(kept, ink, size_pt=16)
    for shape, _box, _text in captions[1:]:
        try:
            delete_shape(shape)
        except Exception:
            continue


def shrink_short_copy_wells(slide) -> None:
    """Collapse a 2″ statement well when it only holds a short title."""
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text or is_layout_chrome(text):
            continue
        height = int(shape.height or 0)
        if height < int(1.4 * EMU_INCH):
            continue
        if len(text.split()) > 10 or "\n" in text:
            continue
        try:
            shape.height = int(0.90 * EMU_INCH)
        except Exception:
            continue


def fill_text_slots(slide, title: str, subtitle: str, body_paras: list[str]) -> None:
    """Write source copy into content wells; keep layout chrome; drop leftover dummy."""
    slots = prototype_text_slots(slide)
    if not slots:
        return
    content = [s for s in slots if not is_layout_chrome(s["text"])]
    title_slot = content.pop(0) if content else None
    if title_slot is not None:
        replace_shape_text(title_slot["shape"], title or "")
        if not title:
            delete_shape(title_slot["shape"])
            title_slot = None
        else:
            fit_title_in_slot(title_slot["shape"], title)
    body_texts: list[str] = [
        p
        for p in body_paras
        if p
        and p != (title or "").strip()
        and p != (subtitle or "").strip()
        and not is_org_chrome(p)
    ]
    to_delete: list[Any] = []
    lead, body_texts = _peel_lead_in(body_texts)
    extra_lines = [line for line in (subtitle, lead) if line]
    columns_preview = _cluster_columns(content)
    if extra_lines:
        wide = _wide_slot(content)
        if wide is not None:
            replace_shape_text(wide["shape"], "\n".join(extra_lines))
            content = [s for s in content if s is not wide]
        elif len(columns_preview) >= 4:
            # Sequence KPI tiles: never mash a full-width intro into card 1.
            pass
        else:
            body_texts = extra_lines + body_texts
    columns = _cluster_columns(content)
    if len(columns) >= 2 and len(body_texts) >= 2:
        chunks = _split_into(body_texts, len(columns))
        for i, col in enumerate(columns):
            chunk = chunks[i] if i < len(chunks) else []
            to_delete.extend(_fill_column(col, chunk))
    elif _looks_like_comparison_grid(content) and (
        not body_texts or (len(body_texts) == 1 and is_short_caption(body_texts[0]))
    ):
        to_delete.extend(_fill_caption_on_grid(content, body_texts))
    else:
        to_delete.extend(_fill_column(content, body_texts))
    for shape in list(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        if is_leftover_prototype_copy(shape.text_frame.text or ""):
            clear_shape_text(shape)
            to_delete.append(shape)
    for shape in to_delete:
        try:
            delete_shape(shape)
        except Exception:
            continue
    expand_text_wells(slide)


def stack_copy_above_photos(slide, tokens: dict[str, Any]) -> None:
    """Keep photos in front of empty cards, but never covering source copy."""
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    slide_area = max(1, slide_w * slide_h)
    pictures = []
    for shape in list(slide.shapes):
        if shape.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            continue
        box = _shape_box(shape)
        if is_brand_lockup(
            box["width"] * box["height"],
            slide_area,
            box["left"],
            box["top"],
            box["width"],
            box["height"],
            slide_w,
            slide_h,
        ):
            continue
        pictures.append(shape)
    for shape in pictures:
        bring_picture_to_front(shape)
    pic_boxes = [_shape_box(s) for s in pictures]
    if not pic_boxes:
        return
    for shape in list(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        box = _shape_box(shape)
        if any(overlap_fraction(box, pic) > 0.08 for pic in pic_boxes):
            el = shape._element
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                parent.append(el)


def strip_pic_locks(element) -> None:
    """Remove picLocks / spLocks so staff can crop, resize, Change Picture, and delete."""
    for ns_tag in (f"{{{A_NS}}}picLocks", f"{{{A_NS}}}spLocks"):
        for locks in list(element.findall(f".//{ns_tag}")):
            parent = locks.getparent()
            if parent is not None:
                parent.remove(locks)


def sanitize_new_picture(shape) -> None:
    """New add_picture pics: no locks, empty cNvPicPr / nvPr (no placeholder)."""
    el = shape._element
    strip_pic_locks(el)
    for src in list(el.findall(f".//{{{A_NS}}}srcRect")):
        parent = src.getparent()
        if parent is not None:
            parent.remove(src)
    cnv = el.find(f".//{{{P_NS}}}cNvPicPr")
    if cnv is not None:
        for child in list(cnv):
            cnv.remove(child)
    nv_pr = el.find(f".//{{{P_NS}}}nvPr")
    if nv_pr is not None:
        for child in list(nv_pr):
            nv_pr.remove(child)


def bring_picture_to_front(shape) -> None:
    """Last in spTree is front-most so clicks hit the photo, not the card behind it."""
    el = shape._element
    parent = el.getparent()
    if parent is None:
        return
    parent.remove(el)
    parent.append(el)


def add_unlocked_picture(
    slide,
    blob: bytes,
    box: dict[str, int],
    img_w: int | None,
    img_h: int | None,
    fill_well: bool = False,
):
    crop = None
    if img_w and img_h and fill_well:
        fit = {
            "left": int(box["left"]),
            "top": int(box["top"]),
            "width": int(box["width"]),
            "height": int(box["height"]),
        }
        crop = cover_crop(int(img_w), int(img_h), fit)
    elif img_w and img_h:
        fit = contain_fit(int(img_w), int(img_h), box)
    else:
        fit = box
    picture = slide.shapes.add_picture(
        io.BytesIO(blob),
        int(fit["left"]),
        int(fit["top"]),
        int(fit["width"]),
        int(fit["height"]),
    )
    sanitize_new_picture(picture)
    if crop and any(crop.values()):
        apply_picture_crop(picture, crop)
    bring_picture_to_front(picture)
    return picture


def _is_picture_placeholder(shape) -> bool:
    try:
        if getattr(shape, "is_placeholder", False):
            name = (enum_name(placeholder_type(shape)) or "").upper()
            if name in {"PICTURE", "MEDIA_CLIP"}:
                return True
    except Exception:
        pass
    nv_pr = shape._element.find(f".//{{{P_NS}}}nvPr")
    if nv_pr is None:
        return False
    ph = nv_pr.find(qn("p:ph"))
    if ph is None:
        return False
    return (ph.get("type") or "").lower() in {"pic", "clipart", "media"}


def content_photo_frames(slide, slide_w: int, slide_h: int) -> list[dict[str, Any]]:
    """Template content photo frames (not the top-left brand lockup)."""
    recs: list[dict[str, Any]] = []
    slide_area = max(1, int(slide_w) * int(slide_h))
    for shape in slide.shapes:
        if shape.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            continue
        left = int(shape.left or 0)
        top = int(shape.top or 0)
        width = int(shape.width or 0)
        height = int(shape.height or 0)
        area = width * height
        recs.append(
            {
                "shape": shape,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "area": area,
                "hash": _picture_hash(shape),
                "placeholder": _is_picture_placeholder(shape),
                "logo": is_brand_lockup(area, slide_area, left, top, width, height, slide_w, slide_h),
            }
        )
    frames: list[dict[str, Any]] = []
    for rec in recs:
        if rec["logo"]:
            continue
        siblings = [other for other in recs if not other["logo"] and abs(other["area"] - rec["area"]) / max(rec["area"], 1) <= 0.08]
        if len(siblings) >= 2:
            frames.append(rec)
            continue
        # Any remaining non-logo picture is sample art / a photo well, including cover heroes.
        frames.append(rec)
    frames.sort(
        key=lambda item: (
            0 if item["placeholder"] else 1,
            -int(item["area"]),
            int(item["left"]),
            int(item["top"]),
        )
    )
    return frames


def _load_picture_blob(placement: dict[str, Any]) -> tuple[bytes, int | None, int | None] | None:
    image = placement.get("image") or {}
    path = image.get("path")
    if not path or not Path(path).exists():
        return None
    blob = Path(path).read_bytes()
    blob, _ext = maybe_recompress(blob, image.get("ext") or "png")
    px = image.get("px") or [None, None]
    img_w, img_h = px[0], px[1]
    if not img_w or not img_h:
        try:
            from PIL import Image

            with Image.open(io.BytesIO(blob)) as im:
                img_w, img_h = im.size
        except Exception:
            img_w, img_h = None, None
    return blob, img_w, img_h


def _normalized_center(box: dict[str, Any], width: int, height: int) -> tuple[float, float]:
    cx = (int(box.get("left") or 0) + int(box.get("width") or 0) / 2) / max(1, width)
    cy = (int(box.get("top") or 0) + int(box.get("height") or 0) / 2) / max(1, height)
    return cx, cy


def _match_photo_frame(placement: dict[str, Any], frames: list[dict[str, Any]], tokens: dict[str, Any]) -> dict[str, Any] | None:
    if not frames:
        return None
    src_w = max(1, int(tokens.get("_src_w") or tokens["slide_width"]))
    src_h = max(1, int(tokens.get("_src_h") or tokens["slide_height"]))
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    bounds = placement.get("bounds") or {}
    if not bounds:
        return frames[0]
    sx, sy = _normalized_center(bounds, src_w, src_h)

    def score(frame: dict[str, Any]) -> tuple[float, int]:
        fx, fy = _normalized_center(frame, slide_w, slide_h)
        return ((sx - fx) ** 2 + (sy - fy) ** 2, -int(frame.get("area") or 0))

    return min(frames, key=score)


def _text_boxes(slide) -> list[dict[str, int]]:
    boxes: list[dict[str, int]] = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        if not (shape.text_frame.text or "").strip():
            continue
        boxes.append(_shape_box(shape))
    return boxes


def _free_photo_box(slide, tokens: dict[str, Any], img_w: int | None, img_h: int | None) -> dict[str, int]:
    """A photo well that does not sit on top of prototype copy."""
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    safe = content_safe_box(tokens, tokens.get("_layout_index"))
    texts = _text_boxes(slide)
    candidates = [
        {"left": int(slide_w * 0.50), "top": 0, "width": int(slide_w * 0.50), "height": slide_h},
        {"left": safe["left"], "top": int(slide_h * 0.68), "width": safe["width"], "height": int(slide_h * 0.28)},
        {
            "left": int(safe["left"] + safe["width"] * 0.55),
            "top": int(safe["top"] + safe["height"] * 0.15),
            "width": int(safe["width"] * 0.45),
            "height": int(safe["height"] * 0.80),
        },
    ]
    pics = [
        _shape_box(shape)
        for shape in slide.shapes
        if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}
    ]
    for cand in candidates:
        box = clamp_box_to_slide(cand, slide_w, slide_h)
        if all(overlap_fraction(box, text) < 0.20 for text in texts) and all(
            overlap_fraction(box, pic) < 0.20 for pic in pics
        ):
            return box
    # Last resort: modest lower-right card, not the full content-safe area.
    return clamp_box_to_slide(
        {
            "left": int(slide_w * 0.62),
            "top": int(slide_h * 0.55),
            "width": int(slide_w * 0.32),
            "height": int(slide_h * 0.38),
        },
        slide_w,
        slide_h,
    )


def _placement_is_icon(placement: dict[str, Any], tokens: dict[str, Any] | None = None) -> bool:
    image = dict(placement.get("image") or {})
    rec = {
        **image,
        "bounds": placement.get("bounds") or {},
        "area_ratio": placement.get("area_ratio") or 0,
    }
    if tokens and not rec.get("area_ratio"):
        bounds = rec.get("bounds") or {}
        sw = max(1, int(tokens.get("_src_w") or tokens.get("slide_width") or 1))
        sh = max(1, int(tokens.get("_src_h") or tokens.get("slide_height") or 1))
        rec["area_ratio"] = int(bounds.get("width") or 0) * int(bounds.get("height") or 0) / (sw * sh)
    if is_decorative_icon(rec):
        return True
    loaded = _load_picture_blob(placement)
    if loaded is None:
        return False
    blob, img_w, img_h = loaded
    if rec.get("area_ratio") and float(rec["area_ratio"]) >= 0.08:
        return False
    return bool(img_w and img_h and max(int(img_w), int(img_h)) <= 256)


def _dense_text_prototype(tokens: dict[str, Any]) -> bool:
    idx = tokens.get("_layout_index")
    proto = next(
        (p for p in tokens.get("prototypes") or [] if int(p.get("index") or -1) == int(idx or -1)),
        None,
    )
    blob = f"{(proto or {}).get('name') or ''} {(proto or {}).get('notes') or ''}".lower()
    return any(
        token in blob
        for token in (
            "sequence",
            "three parts",
            "two points",
            "contents",
            "people",
            "kicker",
            "the point",
            "by the numbers",
            "cover layout",
        )
    )


def place_icon_pictures(slide, placements: list[dict[str, Any]], tokens: dict[str, Any]) -> None:
    """Never stamp UI glyphs onto Sequence/card frames or leftover photo wells.

    Decorative 256px icons are dropped (drop_icon) so they cannot sit on copy.
    """
    if _dense_text_prototype(tokens):
        return
    icons: list[tuple[bytes, int | None, int | None]] = []
    for placement in placements:
        loaded = _load_picture_blob(placement)
        if loaded is None:
            continue
        icons.append(loaded)
    if not icons:
        return
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    size = int(0.7 * EMU_INCH)
    gap = int(0.12 * EMU_INCH)
    texts = _text_boxes(slide)
    safe = content_safe_box(tokens, tokens.get("_layout_index"))
    x = int(safe["left"])
    y = min(int(safe["top"] + safe["height"] - size), int(slide_h - size - 0.35 * EMU_INCH))
    for blob, img_w, img_h in icons:
        box = clamp_box_to_slide(
            {"left": x, "top": y, "width": size, "height": size},
            slide_w,
            slide_h,
        )
        if any(overlap_fraction(box, text) > 0.12 for text in texts):
            x += size + gap
            continue
        add_unlocked_picture(slide, blob, box, img_w, img_h)
        x += size + gap


def purge_empty_table_chrome(slide, tokens: dict[str, Any]) -> None:
    """Drop comparison row boards and hairlines once a live table/chart will sit there."""
    slide_w = int(tokens["slide_width"])
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        if (shape.text_frame.text or "").strip():
            continue
        box = _shape_box(shape)
        wide = box["width"] >= int(slide_w * 0.55)
        row_bar = wide and int(0.4 * EMU_INCH) <= box["height"] <= int(1.25 * EMU_INCH)
        hairline = box["height"] <= int(0.06 * EMU_INCH) and box["width"] >= int(slide_w * 0.12)
        if row_bar or hairline:
            try:
                delete_shape(shape)
            except Exception:
                continue


def promote_text_into_empty_cards(slide, tokens: dict[str, Any]) -> None:
    """When a photo well was deleted, move column copy up into the empty card."""
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    pic_boxes = []
    for shape in slide.shapes:
        if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            pic_boxes.append(_shape_box(shape))
    cards: list[tuple[Any, dict[str, int]]] = []
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        if (shape.text_frame.text or "").strip():
            continue
        box = _shape_box(shape)
        if box["width"] < int(2.0 * EMU_INCH) or box["height"] < int(2.0 * EMU_INCH):
            continue
        if box["height"] < int(slide_h * 0.22):
            continue
        if any(overlap_fraction(box, pic) > 0.20 for pic in pic_boxes):
            continue
        cards.append((shape, box))
    pad = int(0.28 * EMU_INCH)
    gap = int(0.10 * EMU_INCH)
    for _card, box in cards:
        nested = []
        for shape in list(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            text = (shape.text_frame.text or "").strip()
            if not text or is_layout_chrome(text):
                continue
            other = _shape_box(shape)
            cx = other["left"] + other["width"] / 2
            cy = other["top"] + other["height"] / 2
            if box["left"] <= cx <= box["left"] + box["width"] and box["top"] <= cy <= box["top"] + box["height"]:
                nested.append(shape)
        if not nested:
            continue
        nested.sort(key=lambda s: int(s.top or 0))
        y = box["top"] + pad
        inner_h = box["height"] - 2 * pad
        share = max(int(0.5 * EMU_INCH), inner_h // max(1, len(nested)))
        for shape in nested:
            try:
                shape.top = int(y)
                shape.left = int(box["left"] + pad)
                shape.width = int(max(int(0.8 * EMU_INCH), box["width"] - 2 * pad))
                shape.height = int(share)
            except Exception:
                continue
            y += int(shape.height or share) + gap


def purge_rules_over_text(slide, tokens: dict[str, Any]) -> None:
    """Drop prototype lime hairlines that cut through source type."""
    slide_w = int(tokens["slide_width"])
    texts = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text or is_layout_chrome(text):
            continue
        texts.append(_shape_box(shape))
    if not texts:
        return
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if getattr(shape, "has_text_frame", False) and (shape.text_frame.text or "").strip():
            continue
        box = _shape_box(shape)
        hairline = box["height"] <= int(0.08 * EMU_INCH) and box["width"] >= int(slide_w * 0.08)
        if not hairline:
            continue
        cuts = False
        for text in texts:
            y_hit = text["top"] < box["top"] < text["top"] + text["height"]
            x_hit = not (box["left"] + box["width"] < text["left"] or box["left"] > text["left"] + text["width"])
            if y_hit and x_hit:
                cuts = True
                break
            if overlap_fraction(box, text) > 0.02:
                cuts = True
                break
        if cuts:
            try:
                delete_shape(shape)
            except Exception:
                continue


def collapse_overlapping_text(slide) -> None:
    """Delete leftover copy layers that sit on top of the same card well."""
    items: list[tuple[Any, dict[str, int], str]] = []
    for shape in list(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text or is_layout_chrome(text):
            continue
        items.append((shape, _shape_box(shape), text))
    deleted: set[int] = set()
    for i, (shape_a, box_a, text_a) in enumerate(items):
        if id(shape_a) in deleted:
            continue
        for shape_b, box_b, text_b in items[i + 1 :]:
            if id(shape_b) in deleted:
                continue
            frac = max(overlap_fraction(box_a, box_b), overlap_fraction(box_b, box_a))
            if frac < 0.12:
                continue
            area_a = max(1, box_a["width"] * box_a["height"])
            area_b = max(1, box_b["width"] * box_b["height"])
            keep, drop, keep_text, drop_text = (
                (shape_a, shape_b, text_a, text_b) if area_a >= area_b else (shape_b, shape_a, text_b, text_a)
            )
            if drop_text and drop_text not in keep_text and keep_text not in drop_text:
                replace_shape_text(keep, keep_text + "\n" + drop_text)
            try:
                clear_shape_text(drop)
                delete_shape(drop)
            except Exception:
                continue
            deleted.add(id(drop))


COVER_CONTAIN_MAX_LEFTOVER = 0.28


def _is_cover_photo_slide(role: str | None, frames: list[dict[str, Any]]) -> bool:
    """Only the Presentation Title slide uses the keep-or-contain collage rule."""
    return role == "TITLE" and bool(frames)


def _place_cover_photos(
    slide,
    sources: list[dict[str, Any]],
    frames: list[dict[str, Any]],
    tokens: dict[str, Any],
    flags: list[str],
) -> None:
    """Contain-fit one photo in the hero well, or keep the official title collage.

    A cover mosaic (several tiles, or one image that would letterbox) looks worse
    than the template aerial/collage — keep that stock (PR #9) and never overflow
    a second collage onto the cream title panel.
    """
    if not frames:
        return
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    hero = max(frames, key=lambda item: int(item.get("area") or 0))
    box = clamp_box_to_slide(
        {"left": hero["left"], "top": hero["top"], "width": hero["width"], "height": hero["height"]},
        slide_w,
        slide_h,
    )
    loaded_sources: list[tuple[dict[str, Any], bytes, int, int]] = []
    for placement in sources:
        loaded = _load_picture_blob(placement)
        if loaded is None:
            continue
        blob, img_w, img_h = loaded
        if not img_w or not img_h:
            continue
        loaded_sources.append((placement, blob, int(img_w), int(img_h)))
    if len(loaded_sources) != 1:
        return
    _placement, blob, img_w, img_h = loaded_sources[0]
    leftover = contain_leftover_fraction(img_w, img_h, box)
    if leftover > COVER_CONTAIN_MAX_LEFTOVER:
        flags.append("COVER_KEPT_TEMPLATE")
        return
    delete_shape(hero["shape"])
    add_unlocked_picture(slide, blob, box, img_w, img_h, fill_well=False)


def place_source_photos(
    slide,
    placements: list[dict[str, Any]],
    tokens: dict[str, Any],
    flags: list[str],
    role: str | None = None,
) -> None:
    """Swap a well only when a real source photograph maps into it; keep unused stock art."""
    sources: list[dict[str, Any]] = []
    for placement in placements:
        if _load_picture_blob(placement) is None:
            continue
        if _placement_is_icon(placement, tokens):
            continue
        sources.append(placement)
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    frames = content_photo_frames(slide, slide_w, slide_h)
    if _is_cover_photo_slide(role, frames):
        _place_cover_photos(slide, sources, frames, tokens, flags)
        return
    used_ids: set[int] = set()
    for placement in sources:
        loaded = _load_picture_blob(placement)
        if loaded is None:
            continue
        blob, img_w, img_h = loaded
        remaining = [frame for frame in frames if id(frame["shape"]) not in used_ids]
        frame = _match_photo_frame(placement, remaining, tokens)
        if frame is not None:
            box = clamp_box_to_slide(
                {"left": frame["left"], "top": frame["top"], "width": frame["width"], "height": frame["height"]},
                slide_w,
                slide_h,
            )
            delete_shape(frame["shape"])
            used_ids.add(id(frame["shape"]))
            add_unlocked_picture(
                slide,
                blob,
                box,
                img_w,
                img_h,
                fill_well=not is_banner_photo_well(box, slide_w, slide_h),
            )
            continue
        add_unlocked_picture(slide, blob, _free_photo_box(slide, tokens, img_w, img_h), img_w, img_h)
        flags.append("PICTURE_OVERFLOW")


def replace_largest_picture(slide, image_path: str, slide_w: int, slide_h: int) -> bool:
    """Back-compat: one source image into the largest content photo frame."""
    frames = content_photo_frames(slide, slide_w, slide_h)
    if not frames:
        return False
    blob = Path(image_path).read_bytes()
    img_w = img_h = None
    try:
        from PIL import Image

        with Image.open(io.BytesIO(blob)) as im:
            img_w, img_h = im.size
    except Exception:
        pass
    frame = frames[0]
    box = {"left": frame["left"], "top": frame["top"], "width": frame["width"], "height": frame["height"]}
    delete_shape(frame["shape"])
    add_unlocked_picture(slide, blob, box, img_w, img_h)
    return True
