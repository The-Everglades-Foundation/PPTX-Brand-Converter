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

from helpers import (
    A_NS,
    P_NS,
    R_NS,
    blob_hash,
    classify_layout_role,
    clamp_box_to_slide,
    contain_fit,
    content_safe_box,
    delete_shape,
    enum_name,
    hex_color,
    is_brand_lockup,
    iter_shapes_abs,
    maybe_recompress,
    overlap_fraction,
    placeholder_type,
    resolve_layout_map,
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


def replace_shape_text(shape, text: str) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.word_wrap = True
    first = tf.paragraphs[0]
    if first.runs:
        first.runs[0].text = text
        for run in first.runs[1:]:
            run.text = ""
    else:
        first.text = text
    for para in tf.paragraphs[1:]:
        for run in para.runs:
            run.text = ""


def fill_text_slots(slide, title: str, subtitle: str, body_paras: list[str]) -> None:
    slots = prototype_text_slots(slide)
    if not slots:
        return
    # Kicker = short uppercase first line; title = next; rest = body.
    kicker = None
    rest = list(slots)
    first = rest[0]
    if len(first["text"]) <= 24 and first["text"] == first["text"].upper() and not first["text"].isdigit():
        kicker = rest.pop(0)
    title_slot = rest.pop(0) if rest else None
    if title_slot is None and kicker is not None:
        title_slot = kicker
        kicker = None
    if title_slot is not None and title:
        replace_shape_text(title_slot["shape"], title)
    elif title_slot is not None:
        replace_shape_text(title_slot["shape"], title_slot["text"] if not title else title)
    body_texts = []
    if subtitle:
        body_texts.append(subtitle)
    body_texts.extend([p for p in body_paras if p])
    # TWO_CONTENT: split remaining slots by x midpoint
    if len(rest) >= 2 and len(body_texts) >= 2:
        mid = (min(s["left"] for s in rest) + max(s["left"] + s["width"] for s in rest)) / 2
        left_slots = [s for s in rest if s["left"] + s["width"] / 2 < mid]
        right_slots = [s for s in rest if s not in left_slots]
        half = max(1, len(body_texts) // 2)
        _fill_slot_group(left_slots, body_texts[:half])
        _fill_slot_group(right_slots, body_texts[half:])
        return
    _fill_slot_group(rest, body_texts)


def _fill_slot_group(slots: list[dict[str, Any]], texts: list[str]) -> None:
    if not slots:
        return
    if not texts:
        # Leave sample copy rather than emptying chrome labels like 01 / COLUMN ONE
        return
    for i, slot in enumerate(slots):
        if i < len(texts):
            value = texts[i]
            if i == len(slots) - 1 and len(texts) > len(slots):
                value = "\n".join(texts[i:])
            replace_shape_text(slot["shape"], value)
        # extra decorative slots keep template copy


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


def add_unlocked_picture(slide, blob: bytes, box: dict[str, int], img_w: int | None, img_h: int | None):
    if img_w and img_h:
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
    for cand in candidates:
        box = clamp_box_to_slide(cand, slide_w, slide_h)
        if all(overlap_fraction(box, text) < 0.20 for text in texts):
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


def purge_sample_photos(slide, tokens: dict[str, Any]) -> None:
    """Drop leftover template sample photographs (never the brand lockup)."""
    sample = set(tokens.get("sample_photo_hashes") or [])
    logos = set(tokens.get("logo_hashes") or [])
    if not sample and not logos:
        return
    for shape in list(slide.shapes):
        if shape.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            continue
        digest = _picture_hash(shape)
        if digest and digest in logos:
            continue
        if digest and digest in sample:
            delete_shape(shape)


def place_source_photos(
    slide,
    placements: list[dict[str, Any]],
    tokens: dict[str, Any],
    flags: list[str],
) -> None:
    """Delete template content frames and add_picture new unlocked pics at that geometry."""
    sources: list[dict[str, Any]] = []
    for placement in placements:
        if _load_picture_blob(placement) is not None:
            sources.append(placement)
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    frames = content_photo_frames(slide, slide_w, slide_h)
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
            add_unlocked_picture(slide, blob, box, img_w, img_h)
            continue
        add_unlocked_picture(slide, blob, _free_photo_box(slide, tokens, img_w, img_h), img_w, img_h)
        flags.append("PICTURE_OVERFLOW")
    # Always drop leftover sample art, including slides with no source photos.
    for frame in frames:
        if id(frame["shape"]) in used_ids:
            continue
        try:
            delete_shape(frame["shape"])
        except Exception:
            continue
    purge_sample_photos(slide, tokens)


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
    for extra in frames[1:]:
        delete_shape(extra["shape"])
    return True
