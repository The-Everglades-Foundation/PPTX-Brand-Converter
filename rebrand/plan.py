#!/usr/bin/env python3
"""Build plan.json from brand_tokens.json + source_manifest.json."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from helpers import (  # noqa: E402
    content_safe_box,
    dump_json,
    fits,
    layout_index_for_role,
    load_json,
    reduce_body_size,
    reduce_title_size,
)
from colorways import apply_colorway  # noqa: E402


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
    blocks = []
    others = []
    for shape in content_shapes(slide):
        if shape.get("id") == title_id:
            continue
        ph = (shape.get("placeholder_type") or "").upper()
        if ph in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE"}:
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
    min_row = int(0.28 * 914400)
    available = max(min_row, int(safe.get("height") or min_row))
    # header + body rows
    per = max(min_row, int(available / max(2, min(len(rows), 8))))
    capacity = max(1, available // per - 1)
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
    for table_pl in table_pls:
        table = table_pl.get("table") or {}
        chunks = table_row_chunks(table, safe)
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
        if shape.get("kind") == "picture" and not shape.get("drop_logo") and not shape.get("drop_footer")
    ]


def pick_prototype_layout(tokens: dict[str, Any], role: str, pictures: list[dict[str, Any]]) -> int:
    """Prefer a sample slide whose photo wells match the source pictures."""
    default = layout_index_for_role(tokens, role)
    prototypes = tokens.get("prototypes") or []
    if not prototypes:
        return default
    n_pics = len(pictures)
    if not n_pics:
        return default

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
            layout_index = pick_prototype_layout(tokens, role, _source_pictures(slide))
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
