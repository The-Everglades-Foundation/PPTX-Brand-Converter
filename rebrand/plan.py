#!/usr/bin/env python3
"""Build plan.json from brand_tokens.json + source_manifest.json."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from helpers import (  # noqa: E402
    content_safe_box,
    data_content_box,
    dump_json,
    fits,
    is_decorative_icon,
    is_org_chrome,
    layout_index_for_role,
    load_json,
    reduce_body_size,
    reduce_title_size,
)
from colorways import apply_colorway, is_short_caption  # noqa: E402


def _ph_type_name(ph: dict[str, Any]) -> str:
    return (ph.get("type_name") or "").upper()


def layout_placeholders(tokens: dict[str, Any], layout_index: int) -> list[dict[str, Any]]:
    layouts = tokens.get("layouts") or []
    for layout in layouts:
        if int(layout["index"]) == int(layout_index):
            return list(layout.get("placeholders") or [])
    return []


def placeholder_box(placeholders: list[dict[str, Any]], kinds: set[str]) -> dict[str, int] | None:
    matches = [ph for ph in placeholders if _ph_type_name(ph) in kinds]
    if not matches:
        return None
    matches.sort(key=lambda ph: int(ph.get("left") or 0))
    ph = matches[0]
    return {
        "left": int(ph["left"]),
        "top": int(ph["top"]),
        "width": int(ph["width"]),
        "height": int(ph["height"]),
        "idx": ph.get("idx"),
        "name": ph.get("name"),
        "type_name": ph.get("type_name"),
    }


def body_placeholders_left_to_right(placeholders: list[dict[str, Any]]) -> list[dict[str, Any]]:
    matches = [ph for ph in placeholders if _ph_type_name(ph) in {"BODY", "OBJECT", "VERTICAL_BODY"}]
    matches.sort(key=lambda ph: int(ph.get("left") or 0))
    return matches


def content_shapes(slide: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        s
        for s in slide.get("shapes") or []
        if s.get("kind") != "group"
        and not s.get("drop_footer")
        and not s.get("drop_logo")
        and not s.get("drop_icon")
    ]


def title_shape(slide: dict[str, Any]) -> dict[str, Any] | None:
    for shape in slide.get("shapes") or []:
        ph = (shape.get("placeholder_type") or "").upper()
        if ph in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE"} and (shape.get("text") or "").strip():
            return shape
    title = (slide.get("title_text") or "").strip()
    if not title:
        return None
    for shape in slide.get("shapes") or []:
        if (shape.get("text") or "").strip() == title:
            return shape
    return None


def split_text_blocks(slide: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    title = title_shape(slide)
    title_id = title.get("id") if title else None
    title_text = (slide.get("title_text") or "").strip()
    subtitle_text = (slide.get("subtitle_text") or "").strip()
    blocks = []
    others = []
    for shape in content_shapes(slide):
        if shape.get("id") == title_id:
            continue
        ph = (shape.get("placeholder_type") or "").upper()
        if ph in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE", "SUBTITLE"}:
            continue
        text = (shape.get("text") or "").strip()
        if text and text in {title_text, subtitle_text}:
            continue
        if shape.get("kind") == "text" and shape.get("text_frame"):
            blocks.append(shape)
        else:
            others.append(shape)
    return blocks, others


def estimate_title_split(title: str, box: dict[str, int] | None, start_size: int) -> tuple[int, bool]:
    if not box:
        return start_size, False
    size = reduce_title_size(title or "", box["width"], box["height"], start_size)
    needs_split = not fits([title or ""], size, box["width"], box["height"], line_spacing=1.1)
    return size, needs_split


def paragraph_chunks(paragraphs: list[dict[str, Any]], box: dict[str, int] | None) -> tuple[list[list[dict[str, Any]]], int]:
    if not paragraphs:
        return [[]], 18
    if not box:
        return [paragraphs], 18
    size = reduce_body_size(paragraphs, box["width"], box["height"])
    if fits(paragraphs, size, box["width"], box["height"]):
        return [paragraphs], size
    # Split on a paragraph boundary near the midpoint.
    mid = max(1, len(paragraphs) // 2)
    return [paragraphs[:mid], paragraphs[mid:]], size


def table_row_chunks(table: dict[str, Any], safe: dict[str, int], font_pt: int = 11) -> list[dict[str, Any]]:
    rows = table.get("cells") or []
    if not rows:
        return []
    header = rows[0]
    body = rows[1:] if len(rows) > 1 else []
    min_row = int(0.32 * 914400)
    available = max(min_row, int(safe.get("height") or min_row))
    # header occupies one row; remaining height is body rows that actually fit.
    capacity = max(1, available // min_row - 1)
    if not body:
        return [{"header": header, "rows": [], "font_pt": font_pt, "cont": False}]
    chunks = []
    for i in range(0, len(body), capacity):
        chunk_rows = body[i : i + capacity]
        chunks.append({"header": header, "rows": chunk_rows, "font_pt": font_pt, "cont": i > 0})
    return chunks or [{"header": header, "rows": body, "font_pt": font_pt, "cont": False}]


def placements_for_slide(
    slide: dict[str, Any],
    role: str,
    tokens: dict[str, Any],
    layout_index: int,
) -> list[dict[str, Any]]:
    placeholders = layout_placeholders(tokens, layout_index)
    title_box = placeholder_box(placeholders, {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE"})
    subtitle_box = placeholder_box(placeholders, {"SUBTITLE"})
    bodies = body_placeholders_left_to_right(placeholders)
    picture_box = placeholder_box(placeholders, {"PICTURE", "MEDIA_CLIP"})
    placements: list[dict[str, Any]] = []
    title_src = title_shape(slide)
    if slide.get("title_text"):
        placements.append(
            {
                "kind": "title",
                "shape_id": title_src.get("id") if title_src else None,
                "target": "placeholder",
                "placeholder": title_box,
                "text": slide.get("title_text"),
            }
        )
    if slide.get("subtitle_text"):
        placements.append(
            {
                "kind": "subtitle",
                "target": "placeholder" if subtitle_box else "textbox",
                "placeholder": subtitle_box,
                "text": slide.get("subtitle_text"),
            }
        )

    text_blocks, other = split_text_blocks(slide)
    # Drop the title text block if we already placed it as title.
    title_text = (slide.get("title_text") or "").strip()
    text_blocks = [b for b in text_blocks if (b.get("text") or "").strip() != title_text]

    if role == "TWO_CONTENT" and len(bodies) >= 2:
        mid = int(tokens["slide_width"]) / 2
        left_blocks = [b for b in text_blocks if b["bounds"]["left"] + b["bounds"]["width"] / 2 < mid]
        right_blocks = [b for b in text_blocks if b not in left_blocks]
        if not left_blocks and not right_blocks and text_blocks:
            half = max(1, len(text_blocks) // 2)
            left_blocks, right_blocks = text_blocks[:half], text_blocks[half:]
        for target, blocks in ((bodies[0], left_blocks), (bodies[1], right_blocks)):
            paras = []
            for block in blocks:
                paras.extend((block.get("text_frame") or {}).get("paragraphs") or [])
            placements.append(
                {
                    "kind": "body",
                    "target": "placeholder",
                    "placeholder": target,
                    "paragraphs": paras,
                    "shape_ids": [b.get("id") for b in blocks],
                }
            )
        text_blocks = []
    elif text_blocks and bodies:
        paras = []
        for block in text_blocks:
            paras.extend((block.get("text_frame") or {}).get("paragraphs") or [])
        placements.append(
            {
                "kind": "body",
                "target": "placeholder",
                "placeholder": bodies[0],
                "paragraphs": paras,
                "shape_ids": [b.get("id") for b in text_blocks],
            }
        )
        text_blocks = []
    elif text_blocks and subtitle_box and role == "TITLE":
        paras = []
        for block in text_blocks:
            paras.extend((block.get("text_frame") or {}).get("paragraphs") or [])
        placements.append(
            {
                "kind": "subtitle",
                "target": "placeholder",
                "placeholder": subtitle_box,
                "paragraphs": paras,
            }
        )
        text_blocks = []

    for block in text_blocks:
        kind = "caption" if (block.get("area_ratio") or 0) < 0.08 else "textbox"
        placements.append(
            {
                "kind": kind,
                "target": "textbox",
                "shape_id": block.get("id"),
                "bounds": block.get("bounds"),
                "paragraphs": (block.get("text_frame") or {}).get("paragraphs") or [],
            }
        )

    for shape in other:
        kind = shape.get("kind")
        if kind == "picture":
            target = "picture_placeholder" if role == "PICTURE" and picture_box else "picture"
            placements.append(
                {
                    "kind": "picture",
                    "target": target,
                    "placeholder": picture_box if target == "picture_placeholder" else None,
                    "shape_id": shape.get("id"),
                    "bounds": shape.get("bounds"),
                    "image": shape.get("image"),
                    "area_ratio": shape.get("area_ratio"),
                }
            )
        elif kind == "table":
            placements.append(
                {
                    "kind": "table",
                    "target": "content_safe",
                    "shape_id": shape.get("id"),
                    "bounds": shape.get("bounds"),
                    "table": shape.get("table"),
                }
            )
        elif kind == "chart":
            placements.append(
                {
                    "kind": "chart",
                    "target": "content_safe",
                    "shape_id": shape.get("id"),
                    "bounds": shape.get("bounds"),
                    "chart": shape.get("chart"),
                }
            )
        elif kind == "smartart":
            paras = shape.get("smartart_text") or (shape.get("text_frame") or {}).get("paragraphs") or []
            placements.append(
                {
                    "kind": "smartart",
                    "target": "placeholder" if bodies else "textbox",
                    "placeholder": bodies[0] if bodies else None,
                    "paragraphs": paras,
                    "flag": "SMARTART_FLATTENED",
                }
            )
        elif kind == "video":
            placements.append(
                {
                    "kind": "media",
                    "target": "content_safe",
                    "shape_id": shape.get("id"),
                    "bounds": shape.get("bounds"),
                    "flag": "MEDIA_MISSING",
                    "name": shape.get("name"),
                }
            )
        elif kind == "text":
            if _callout_worth_keeping(shape):
                placements.append(
                    {
                        "kind": "callout",
                        "target": "autoshape",
                        "shape_id": shape.get("id"),
                        "bounds": shape.get("bounds"),
                        "paragraphs": (shape.get("text_frame") or {}).get("paragraphs") or [],
                    }
                )
        elif shape.get("is_line"):
            placements.append(
                {
                    "kind": "line",
                    "target": "connector",
                    "shape_id": shape.get("id"),
                    "bounds": shape.get("bounds"),
                }
            )
        elif kind == "text" or (kind == "other" and (shape.get("text") or "").strip()):
            if _callout_worth_keeping(shape):
                placements.append(
                    {
                        "kind": "callout",
                        "target": "autoshape",
                        "shape_id": shape.get("id"),
                        "bounds": shape.get("bounds"),
                        "paragraphs": (shape.get("text_frame") or {}).get("paragraphs") or [],
                    }
                )
    return placements


MIN_CALLOUT_EMU = int(0.08 * 914400)


def _callout_worth_keeping(shape: dict[str, Any]) -> bool:
    """Skip empty PptxGenJS cards / hairlines that become lime overlay rectangles."""
    text = (shape.get("text") or "").strip()
    if not text:
        return False
    bounds = shape.get("bounds") or {}
    if int(bounds.get("width") or 0) < MIN_CALLOUT_EMU:
        return False
    if int(bounds.get("height") or 0) < MIN_CALLOUT_EMU:
        return False
    return True


def split_entries_for(entry: dict[str, Any], tokens: dict[str, Any]) -> list[dict[str, Any]]:
    """Insert overflow / table continuation entries when content cannot fit."""
    layout_index = int(entry["template_layout_index"])
    placeholders = layout_placeholders(tokens, layout_index)
    title_box = placeholder_box(placeholders, {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE"})
    bodies = body_placeholders_left_to_right(placeholders)
    body_box = None
    if bodies:
        ph = bodies[0]
        body_box = {"width": int(ph["width"]), "height": int(ph["height"])}
    safe = content_safe_box(tokens, layout_index)
    extras: list[dict[str, Any]] = []

    title_pl = next((p for p in entry["placements"] if p["kind"] == "title"), None)
    if title_pl and title_box:
        start = 40 if entry["role"] == "TITLE" else (36 if entry["role"] == "SECTION" else 28)
        size, needs = estimate_title_split(title_pl.get("text") or "", title_box, start)
        title_pl["font_pt"] = size
        if needs:
            text = title_pl.get("text") or ""
            words = text.split()
            mid = max(1, len(words) // 2)
            title_pl["text"] = " ".join(words[:mid])
            clone = json.loads(json.dumps(entry))
            clone["placements"] = [p for p in clone["placements"] if p["kind"] != "title"]
            clone["placements"].insert(0, {**title_pl, "text": " ".join(words[mid:])})
            extras.append(clone)

    body_pls = [p for p in entry["placements"] if p["kind"] == "body"]
    new_body_entries = []
    kept_body = []
    for body in body_pls:
        box = body.get("placeholder") or body_box or safe
        chunks, size = paragraph_chunks(body.get("paragraphs") or [], box)
        body["font_pt"] = size
        if len(chunks) == 1:
            body["paragraphs"] = chunks[0]
            kept_body.append(body)
            continue
        body["paragraphs"] = chunks[0]
        kept_body.append(body)
        for chunk in chunks[1:]:
            clone = json.loads(json.dumps(entry))
            clone["placements"] = [p for p in clone["placements"] if p["kind"] not in {"body", "table", "chart", "picture"}]
            clone["placements"].append({**body, "paragraphs": chunk})
            new_body_entries.append(clone)
    entry["placements"] = [p for p in entry["placements"] if p["kind"] != "body"] + kept_body
    extras.extend(new_body_entries)

    table_pls = [p for p in entry["placements"] if p["kind"] == "table"]
    new_tables = []
    kept_tables = []
    table_well = data_content_box(tokens, layout_index)
    table_safe = dict(table_well)
    table_safe["height"] = max(int(0.28 * 914400), int(table_well["height"]) - int(0.55 * 914400))
    for table_pl in table_pls:
        table = table_pl.get("table") or {}
        chunks = table_row_chunks(table, table_safe)
        if len(chunks) <= 1:
            if chunks:
                table_pl["table_chunk"] = chunks[0]
            kept_tables.append(table_pl)
            continue
        table_pl["table_chunk"] = chunks[0]
        kept_tables.append(table_pl)
        for chunk in chunks[1:]:
            clone = json.loads(json.dumps(entry))
            clone["placements"] = [p for p in clone["placements"] if p["kind"] != "table"]
            clone["placements"].append({**table_pl, "table_chunk": chunk})
            new_tables.append(clone)
    entry["placements"] = [p for p in entry["placements"] if p["kind"] != "table"] + kept_tables
    extras.extend(new_tables)

    if not extras:
        return [entry]
    all_entries = [entry, *extras]
    total = len(all_entries)
    titled = []
    for i, item in enumerate(all_entries, start=1):
        item = json.loads(json.dumps(item))
        flags = list(item.get("flags") or [])
        suffix = None
        table_chunk = next((p.get("table_chunk") for p in item["placements"] if p.get("kind") == "table"), None)
        if table_chunk and table_chunk.get("cont"):
            suffix = "(cont.)"
        elif total > 1:
            suffix = f"({i}/{total})"
        title_pl = next((p for p in item["placements"] if p["kind"] == "title"), None)
        if title_pl and suffix:
            base = (title_pl.get("text") or "").split("(")[0].strip()
            title_pl["text"] = f"{base} {suffix}".strip()
        if suffix:
            flags.append("SPLIT")
            item["split"] = {"part": i, "of": total, "title_suffix": suffix}
        item["flags"] = sorted(set(flags))
        titled.append(item)
    return titled


def _source_pictures(slide: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        shape
        for shape in slide.get("shapes") or []
        if shape.get("kind") == "picture"
        and not shape.get("drop_logo")
        and not shape.get("drop_footer")
        and not is_decorative_icon(shape)
    ]


def _proto_blob(proto: dict[str, Any]) -> str:
    return f"{proto.get('name') or ''} {proto.get('notes') or ''}".lower()


def prototype_index_matching(tokens: dict[str, Any], *needles: str) -> int | None:
    lowered = [n.lower() for n in needles]
    for proto in tokens.get("prototypes") or []:
        blob = _proto_blob(proto)
        if any(needle in blob for needle in lowered):
            return int(proto["index"])
    return None


def _text_column_count(slide: dict[str, Any] | None) -> int:
    if not slide:
        return 0
    title = title_shape(slide)
    title_id = title.get("id") if title else None
    blocks = []
    for shape in content_shapes(slide):
        if shape.get("id") == title_id:
            continue
        if shape.get("kind") != "text":
            continue
        text = (shape.get("text") or "").strip()
        if not text:
            continue
        ph = (shape.get("placeholder_type") or "").upper()
        if ph in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE", "SUBTITLE"}:
            continue
        blocks.append(shape)
    if len(blocks) < 2:
        return len(blocks)
    widths = [int((s.get("bounds") or {}).get("width") or 0) for s in blocks]
    median = sorted(widths)[len(widths) // 2] if widths else 0
    narrow = [
        s
        for s in blocks
        if int((s.get("bounds") or {}).get("width") or 0) <= max(int(median * 1.6), int(4.2 * 914400))
    ]
    if len(narrow) >= 2:
        blocks = narrow
    ordered = sorted(blocks, key=lambda s: int((s.get("bounds") or {}).get("left") or 0))
    groups = 1
    prev_right = int((ordered[0].get("bounds") or {}).get("left") or 0) + int(
        (ordered[0].get("bounds") or {}).get("width") or 0
    )
    prev_center = int((ordered[0].get("bounds") or {}).get("left") or 0) + int(
        (ordered[0].get("bounds") or {}).get("width") or 0
    ) / 2
    for shape in ordered[1:]:
        bounds = shape.get("bounds") or {}
        left = int(bounds.get("left") or 0)
        width = int(bounds.get("width") or 0)
        center = left + width / 2
        if left >= prev_right - int(0.2 * 914400) and center - prev_center >= 1.2 * 914400:
            groups += 1
            prev_right = left + width
            prev_center = center
        else:
            prev_right = max(prev_right, left + width)
    return groups


def _is_full_width_block(shape: dict[str, Any], blocks: list[dict[str, Any]]) -> bool:
    bounds = shape.get("bounds") or {}
    width = int(bounds.get("width") or 0)
    if not width or not blocks:
        return False
    rights = [
        int((s.get("bounds") or {}).get("left") or 0) + int((s.get("bounds") or {}).get("width") or 0)
        for s in blocks
    ]
    span = max(rights) - min(int((s.get("bounds") or {}).get("left") or 0) for s in blocks)
    return width >= int(span * 0.72) if span else False


_HEADING_LINE_RE = re.compile(
    r"^(?:(?:step|phase)\s+\d+\b|\d+\s*[·.)-])",
    re.I,
)


def _is_point_heading(text: str) -> bool:
    t = " ".join((text or "").split())
    if not t or is_org_chrome(t):
        return False
    if " — " in t or " – " in t:
        label = t.split(" — ")[0].split(" – ")[0]
        return 1 <= len(label.split()) <= 8
    words = t.split()
    if _HEADING_LINE_RE.match(t):
        return True
    if 1 <= len(words) <= 8 and not t.endswith((".", "!", "?")):
        return True
    return False


def _body_point_count(slide: dict[str, Any] | None) -> int:
    """Heading-like lines inside one body well (Today / Step 1 / Label —)."""
    if not slide:
        return 0
    title = (slide.get("title_text") or "").strip()
    numbered = 0
    emdash = 0
    shorts = 0
    for shape in content_shapes(slide):
        if shape.get("kind") != "text":
            continue
        ph = (shape.get("placeholder_type") or "").upper()
        if ph in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE", "SUBTITLE"}:
            continue
        paras = ((shape.get("text_frame") or {}).get("paragraphs") or [])
        lines = [((p.get("text") or "").strip()) for p in paras] if paras else [(shape.get("text") or "").strip()]
        for line in lines:
            if not line or line == title or is_org_chrome(line):
                continue
            compact = " ".join(line.split())
            if _HEADING_LINE_RE.match(compact):
                numbered += 1
            elif " — " in compact or " – " in compact:
                label = compact.split(" — ")[0].split(" – ")[0]
                if 1 <= len(label.split()) <= 8:
                    emdash += 1
            elif _is_point_heading(compact):
                shorts += 1
    if numbered == 4 and shorts >= 3:
        return 4
    if numbered >= 2 and shorts >= 2:
        return numbered + shorts
    if numbered >= 2:
        return numbered
    if emdash >= 2:
        return emdash + shorts
    return numbered + emdash + shorts


def _entry_has_competing_body(entry: dict[str, Any]) -> bool:
    """True when body copy would fight a table/chart for the same well."""
    blobs: list[str] = []
    para_count = 0
    for placement in entry.get("placements") or []:
        if placement.get("kind") not in {"body", "textbox", "caption", "subtitle"}:
            continue
        paras = placement.get("paragraphs") or []
        text = " ".join((p.get("text") or "") for p in paras) or str(placement.get("text") or "")
        if not (text or "").strip():
            continue
        blobs.append(text)
        filled = [p for p in paras if (p.get("text") or "").strip()]
        para_count += max(1, len(filled))
    if len(blobs) >= 2 or para_count > 2:
        return True
    if not blobs:
        return False
    if is_short_caption(blobs[0]):
        return False
    return len(blobs[0].split()) >= 8


def _content_slot_count(slide: dict[str, Any] | None) -> int:
    """Cards the copy wants: KPI tiles, ALL-CAPS labels, or long statements.

    Spatial x-clustering undercounts PptxGenJS 2-col stacks of 3–5 items, so
    Who's Using / Agentic / Takeaways were landing on Two Points.
    """
    if not slide:
        return 0
    title = title_shape(slide)
    title_id = title.get("id") if title else None
    title_text = (slide.get("title_text") or "").strip()
    blocks: list[dict[str, Any]] = []
    for shape in content_shapes(slide):
        if shape.get("id") == title_id:
            continue
        if shape.get("kind") != "text":
            continue
        text = (shape.get("text") or "").strip()
        if not text or text == title_text:
            continue
        ph = (shape.get("placeholder_type") or "").upper()
        if ph in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE", "SUBTITLE"}:
            continue
        blocks.append(shape)
    if not blocks:
        return 0
    n_caps = 0
    n_stmt = 0
    n_kpi = 0
    for shape in blocks:
        if _is_full_width_block(shape, blocks):
            continue
        text = (shape.get("text") or "").strip()
        words = text.split()
        letters = "".join(ch for ch in text if ch.isalpha())
        if letters and text.upper() == text and 1 <= len(words) <= 6:
            n_caps += 1
        elif len(words) >= 12:
            n_stmt += 1
        elif len(words) <= 8:
            n_kpi += 1
    cols = _text_column_count(slide)
    if n_kpi >= 4:
        return max(4, cols)
    if n_stmt >= 4 and cols <= 2:
        return n_stmt
    if n_caps >= 3 or n_stmt == 3:
        return 3
    if n_stmt >= 4:
        return max(3, cols)
    points = _body_point_count(slide)
    return max(cols, points)


def _dense_text_layout(tokens: dict[str, Any], layout_index: int) -> bool:
    proto = next(
        (p for p in tokens.get("prototypes") or [] if int(p.get("index") or -1) == int(layout_index)),
        None,
    )
    if not proto:
        return False
    blob = _proto_blob(proto)
    return any(
        token in blob
        for token in (
            "three parts",
            "two points",
            "people",
            "kicker",
            "contents",
            "by the numbers",
            "cover layout",
            "sequence",
            "statement slide",
            "the point",
        )
    )


def pick_prototype_layout(
    tokens: dict[str, Any],
    role: str,
    pictures: list[dict[str, Any]],
    slide: dict[str, Any] | None = None,
) -> int:
    """Prefer a sample slide whose photo wells match real photographs."""
    default = layout_index_for_role(tokens, role)
    prototypes = tokens.get("prototypes") or []
    if not prototypes:
        return default
    real = [p for p in pictures if not is_decorative_icon(p)]
    n_pics = len(real)
    cols = _text_column_count(slide)
    slots = _content_slot_count(slide)
    n_tables = sum(1 for s in (slide or {}).get("shapes") or [] if s.get("kind") == "table")
    n_charts = sum(1 for s in (slide or {}).get("shapes") or [] if s.get("kind") == "chart")

    if role == "TITLE":
        return prototype_index_matching(tokens, "cover layout") or default
    if role == "CLOSING":
        return prototype_index_matching(tokens, "closing slide") or default
    if role == "SECTION":
        return prototype_index_matching(tokens, "section divider") or default

    if n_tables and n_pics == 0 and n_charts == 0 and slots < 3 and cols < 3:
        found = prototype_index_matching(tokens, "data table", "comparison")
        if found is not None:
            return found

    if not n_pics:
        if slots >= 4 or cols >= 4:
            return (
                prototype_index_matching(tokens, "timeline or process", "sequence")
                or prototype_index_matching(tokens, "contents or any numbered", "contents")
                or default
            )
        if slots >= 5 and cols <= 2:
            return prototype_index_matching(tokens, "contents or any numbered", "contents") or default
        if role == "TWO_CONTENT" or slots >= 2 or cols >= 2:
            if slots >= 3 or cols >= 3:
                return prototype_index_matching(tokens, "three parts", "three parallel") or default
            return prototype_index_matching(tokens, "two-point", "two points") or default
        if role == "TITLE_BODY":
            if cols <= 1:
                return prototype_index_matching(tokens, "statement slide") or default
            return prototype_index_matching(tokens, "contents or any numbered", "contents") or default
        if role == "TITLE_ONLY":
            if n_tables:
                return prototype_index_matching(tokens, "data table", "comparison") or default
            if n_charts:
                return prototype_index_matching(tokens, "statement slide") or default
            return prototype_index_matching(tokens, "statement slide") or default
        return default

    pictures = real

    src_w = max(1, int(tokens.get("_src_w") or tokens.get("slide_width") or 1))
    src_h = max(1, int(tokens.get("_src_h") or tokens.get("slide_height") or 1))
    tpl_w = max(1, int(tokens.get("slide_width") or 1))
    tpl_h = max(1, int(tokens.get("slide_height") or 1))

    def pic_center(shape: dict[str, Any]) -> tuple[float, float]:
        b = shape.get("bounds") or {}
        return (
            (int(b.get("left") or 0) + int(b.get("width") or 0) / 2) / src_w,
            (int(b.get("top") or 0) + int(b.get("height") or 0) / 2) / src_h,
        )

    src_centers = [pic_center(p) for p in pictures]
    related = {
        "TITLE": {"TITLE"},
        "CLOSING": {"CLOSING"},
        "SECTION": {"SECTION", "TITLE_BODY"},
        "PICTURE": {"PICTURE", "TITLE_BODY"},
        "TWO_CONTENT": {"TWO_CONTENT", "TITLE_BODY", "PICTURE"},
        "TITLE_BODY": {"TITLE_BODY", "TWO_CONTENT", "PICTURE"},
        "TITLE_ONLY": {"TITLE_ONLY", "TITLE_BODY", "PICTURE"},
        "BLANK": {"BLANK", "TITLE_BODY", "PICTURE"},
    }.get(role, {role})

    scored: list[tuple[Any, ...]] = []
    for proto in prototypes:
        frames = proto.get("photo_frames") or []
        n_frames = int(proto.get("photo_frame_count") or len(frames))
        proto_role = proto.get("role")
        if proto_role not in related and n_frames == 0:
            continue
        if n_pics and n_frames == 0 and proto_role not in {"TITLE", "CLOSING"}:
            role_penalty = 80
        else:
            role_penalty = 0 if proto_role in related else 8
        if proto_role == "PICTURE" and role != "PICTURE":
            role_penalty += 12
        if frames:
            max_area = max(int(f.get("width") or 0) * int(f.get("height") or 0) for f in frames)
            if max_area >= tpl_w * tpl_h * 0.7 and role != "PICTURE":
                role_penalty += 20
        count_gap = abs(n_frames - n_pics)
        pos_gap = 0.0
        if frames and src_centers:
            for sx, sy in src_centers:
                best = min(
                    (
                        ((sx - (int(f.get("left") or 0) + int(f.get("width") or 0) / 2) / tpl_w) ** 2)
                        + ((sy - (int(f.get("top") or 0) + int(f.get("height") or 0) / 2) / tpl_h) ** 2)
                    )
                    for f in frames
                )
                pos_gap += best
            pos_gap /= len(src_centers)
        scored.append((role_penalty, count_gap, pos_gap, int(proto["index"])))
    if not scored:
        return default
    scored.sort()
    return int(scored[0][3])


def apply_overrides(slides: list[dict[str, Any]], overrides: list[dict[str, Any]] | None) -> None:
    if not overrides:
        return
    by_index = {int(item["source_index"]): item for item in overrides if "source_index" in item}
    for slide in slides:
        ov = by_index.get(int(slide["index"]))
        if not ov:
            continue
        if ov.get("role"):
            slide["role"] = ov["role"]
        if ov.get("template_layout_index") is not None:
            slide["_layout_override"] = int(ov["template_layout_index"])


def build_plan(
    tokens: dict[str, Any],
    manifest: dict[str, Any],
    overrides: list[dict[str, Any]] | None = None,
    colorway: str | None = None,
) -> dict[str, Any]:
    tokens = apply_colorway(tokens, colorway or tokens.get("colorway"))
    tokens["_src_w"] = int(manifest.get("slide_width") or tokens.get("slide_width") or 1)
    tokens["_src_h"] = int(manifest.get("slide_height") or tokens.get("slide_height") or 1)
    slides = list(manifest.get("slides") or [])
    apply_overrides(slides, overrides)
    entries: list[dict[str, Any]] = []
    warnings: list[str] = []
    if not tokens.get("layouts"):
        warnings.append("Template has no layouts.")
    for slide in slides:
        role = slide.get("role") or "TITLE_ONLY"
        if slide.get("_layout_override") is not None:
            layout_index = int(slide["_layout_override"])
        else:
            layout_index = pick_prototype_layout(tokens, role, _source_pictures(slide), slide)
        layout = next((l for l in tokens.get("layouts") or [] if int(l["index"]) == layout_index), None)
        proto = next((p for p in tokens.get("prototypes") or [] if int(p["index"]) == layout_index), None)
        layout_name = (proto or layout or {}).get("name")
        flags = []
        if slide.get("has_transitions_or_animations"):
            flags.append("ANIMATIONS_DROPPED")
        for shape in slide.get("shapes") or []:
            if shape.get("drop_logo"):
                flags.append("LOGO_DROPPED")
            if shape.get("kind") == "smartart":
                flags.append("SMARTART_FLATTENED")
            if shape.get("kind") == "video":
                flags.append("MEDIA_MISSING")
            if shape.get("kind") == "chart":
                chart = shape.get("chart") or {}
                if chart.get("external") or not chart.get("supported"):
                    flags.append("NEEDS_MANUAL_REBUILD")
        entry = {
            "source_index": int(slide["index"]),
            "role": role,
            "template_layout_index": layout_index,
            "template_layout_name": layout_name,
            "hidden": bool(slide.get("hidden")),
            "notes": slide.get("notes_text") or "",
            "placements": placements_for_slide(slide, role, tokens, layout_index),
            "flags": sorted(set(flags)),
        }
        data_pls = [p for p in entry["placements"] if p.get("kind") in {"table", "chart"}]
        if data_pls and _dense_text_layout(tokens, layout_index) and _entry_has_competing_body(entry):
            text_entry = json.loads(json.dumps(entry))
            text_entry["placements"] = [p for p in text_entry["placements"] if p.get("kind") not in {"table", "chart"}]
            data_entry = json.loads(json.dumps(entry))
            has_table = any(p.get("kind") == "table" for p in data_pls)
            data_index = (
                prototype_index_matching(tokens, "data table", "comparison")
                if has_table
                else prototype_index_matching(tokens, "statement slide")
            )
            if data_index is None:
                data_index = layout_index_for_role(tokens, "TITLE_ONLY")
            data_proto = next(
                (p for p in tokens.get("prototypes") or [] if int(p["index"]) == int(data_index)),
                None,
            )
            data_layout = next(
                (item for item in tokens.get("layouts") or [] if int(item["index"]) == int(data_index)),
                None,
            )
            data_entry["role"] = "TITLE_ONLY"
            data_entry["template_layout_index"] = int(data_index)
            data_entry["template_layout_name"] = (data_proto or data_layout or {}).get("name")
            title_pls = [p for p in data_entry["placements"] if p.get("kind") == "title"]
            if title_pls:
                base = (title_pls[0].get("text") or "").split("(")[0].strip()
                title_pls[0]["text"] = f"{base} (cont.)".strip()
            data_entry["placements"] = title_pls + data_pls
            data_flags = list(data_entry.get("flags") or [])
            data_flags.append("SPLIT")
            data_entry["flags"] = sorted(set(data_flags))
            data_entry["split"] = {"part": 2, "of": 2, "title_suffix": "(cont.)"}
            entries.extend(split_entries_for(text_entry, tokens))
            entries.extend(split_entries_for(data_entry, tokens))
        else:
            entries.extend(split_entries_for(entry, tokens))

    plan = {
        "entries": entries,
        "source_count": len(slides),
        "output_count": len(entries),
        "warnings": warnings,
        "layout_map": tokens.get("layout_map"),
        "colorway": tokens.get("colorway"),
        "colorway_label": tokens.get("colorway_label"),
        "build_mode": tokens.get("build_mode"),
    }
    return plan


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create plan.json from inspect artifacts")
    parser.add_argument("--out-dir", default=".")
    parser.add_argument("--overrides", default=None, help="JSON file or inline JSON array of role overrides")
    parser.add_argument("--colorway", default=None)
    args = parser.parse_args(argv)
    out_dir = Path(args.out_dir)
    tokens = apply_colorway(load_json(out_dir / "brand_tokens.json"), args.colorway)
    dump_json(out_dir / "brand_tokens.json", tokens)
    manifest = load_json(out_dir / "source_manifest.json")
    overrides = None
    if args.overrides:
        raw = Path(args.overrides)
        if raw.exists():
            overrides = json.loads(raw.read_text(encoding="utf-8"))
        else:
            overrides = json.loads(args.overrides)
        if isinstance(overrides, dict) and "entries" in overrides:
            overrides = overrides["entries"]
    plan = build_plan(tokens, manifest, overrides)
    dump_json(out_dir / "plan.json", plan)
    print(f"Wrote {out_dir / 'plan.json'} ({plan['output_count']} output slides from {plan['source_count']} source slides)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
