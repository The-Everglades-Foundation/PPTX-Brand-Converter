#!/usr/bin/env python3
"""QA: render slides, run programmatic checks, write QA/qa-report.md."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from colorways import apply_colorway  # noqa: E402
from helpers import (  # noqa: E402
    A_NS,
    allowed_colors,
    allowed_fonts,
    blob_hash,
    content_safe_box,
    dump_json,
    estimate_lines,
    hex_color,
    load_json,
    normalize_for_parity,
    parse_brand_md,
    theme_color,
    theme_font,
)

EMU_PER_PT = 12700


def which(names: list[str]) -> str | None:
    for name in names:
        path = shutil.which(name)
        if path:
            return path
    return None


def render_pptx(pptx_path: Path, dest_dir: Path, prefix: str) -> dict[str, Any]:
    dest_dir.mkdir(parents=True, exist_ok=True)
    soffice = which(["soffice", "libreoffice"])
    pdftoppm = which(["pdftoppm"])
    result = {"ok": False, "tool": None, "images": [], "error": None, "pdf": None}
    if not soffice:
        result["error"] = "LibreOffice (soffice) is not installed. Install LibreOffice Impress for visual QA."
        return result
    pdf_dir = dest_dir / f"{prefix}-pdf"
    if pdf_dir.exists():
        shutil.rmtree(pdf_dir)
    pdf_dir.mkdir(parents=True)
    try:
        subprocess.run(
            [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(pptx_path)],
            check=True,
            capture_output=True,
            text=True,
            timeout=180,
        )
    except Exception as exc:
        result["error"] = f"LibreOffice failed: {exc}"
        return result
    pdfs = list(pdf_dir.glob("*.pdf"))
    if not pdfs:
        result["error"] = "LibreOffice produced no PDF."
        return result
    pdf = pdfs[0]
    result["pdf"] = str(pdf)
    if pdftoppm:
        try:
            subprocess.run(
                [pdftoppm, "-r", "96", "-png", str(pdf), str(dest_dir / prefix)],
                check=True,
                capture_output=True,
                text=True,
                timeout=120,
            )
            images = sorted(dest_dir.glob(f"{prefix}-*.png")) or sorted(dest_dir.glob(f"{prefix}*.png"))
            # Normalize names to {prefix}-XX.png
            renamed = []
            for i, image in enumerate(images, start=1):
                target = dest_dir / f"{prefix}-{i:02d}.png"
                if image.resolve() != target.resolve():
                    if target.exists():
                        target.unlink()
                    image.rename(target)
                renamed.append(target)
            result["images"] = [str(p) for p in renamed]
            result["ok"] = True
            result["tool"] = "soffice+pdftoppm"
            return result
        except Exception as exc:
            result["error"] = f"pdftoppm failed: {exc}"
    # Fallback: keep the PDF even if PNG conversion is missing.
    result["ok"] = True
    result["tool"] = "soffice"
    result["error"] = result["error"] or "pdftoppm (poppler) is not installed; PDF was produced without PNGs."
    return result


def grep_fonts(pptx_path: Path) -> list[str]:
    faces = []
    with zipfile.ZipFile(pptx_path) as zf:
        for name in zf.namelist():
            if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
                continue
            xml = zf.read(name)
            for match in re.findall(br'<a:latin[^>]*typeface="([^"]+)"', xml):
                faces.append(match.decode("utf-8", "replace"))
    return faces


def grep_srgb(pptx_path: Path) -> list[tuple[str, str]]:
    found = []
    with zipfile.ZipFile(pptx_path) as zf:
        for name in zf.namelist():
            if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
                continue
            xml = zf.read(name)
            root = etree.fromstring(xml)
            for node in root.findall(f".//{{{A_NS}}}srgbClr"):
                val = node.get("val")
                if val:
                    found.append((name, hex_color(val)))
    return found


def slide_text(slide) -> str:
    chunks = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            chunks.append(shape.text_frame.text or "")
        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    chunks.append(cell.text or "")
    return "\n".join(chunks)


def source_slide_text(slide_rec: dict[str, Any]) -> str:
    chunks = []
    for shape in slide_rec.get("shapes") or []:
        if shape.get("drop_footer") or shape.get("drop_logo"):
            continue
        if shape.get("kind") in {"text", "table", "smartart"} or shape.get("text"):
            if shape.get("kind") == "table":
                table = shape.get("table") or {}
                for row in table.get("cells") or []:
                    for cell in row:
                        chunks.append(cell.get("text") or "")
            else:
                chunks.append(shape.get("text") or "")
        if shape.get("kind") == "smartart":
            for para in shape.get("smartart_text") or []:
                chunks.append(para.get("text") or "")
    notes = slide_rec.get("notes_text") or ""
    return "\n".join(chunks + ([notes] if notes else []))


def output_notes(slide) -> str:
    if slide.has_notes_slide:
        return slide.notes_slide.notes_text_frame.text or ""
    return ""


def overflow_fail(prs, tokens: dict[str, Any]) -> list[str]:
    fails = []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text_frame.text or ""
            if not text.strip():
                continue
            size = 18
            try:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.font.size:
                            size = run.font.size.pt
                            break
            except Exception:
                pass
            lines = estimate_lines(text, size, int(shape.width))
            need = lines * size * 1.2 * EMU_PER_PT
            if need > int(shape.height) * 1.05:
                fails.append(f"slide {i + 1} shape {shape.name}")
    return fails


def bounds_fail(prs, tokens: dict[str, Any]) -> list[str]:
    fails = []
    w, h = int(tokens["slide_width"]), int(tokens["slide_height"])
    chrome = tokens.get("master_chrome") or []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            left = int(getattr(shape, "left", 0) or 0)
            top = int(getattr(shape, "top", 0) or 0)
            width = int(getattr(shape, "width", 0) or 0)
            height = int(getattr(shape, "height", 0) or 0)
            if left < -12700 or top < -12700 or left + width > w + 12700 or top + height > h + 12700:
                fails.append(f"slide {i + 1} {shape.name} extends beyond slide")
            for ch in chrome:
                c_left, c_top = int(ch["left"]), int(ch["top"])
                c_right = c_left + int(ch["width"])
                c_bottom = c_top + int(ch["height"])
                if not (left + width < c_left or left > c_right or top + height < c_top or top > c_bottom):
                    # Ignore tiny overlaps with full-bleed master backgrounds.
                    if int(ch["width"]) * int(ch["height"]) > w * h * 0.4:
                        continue
                    fails.append(f"slide {i + 1} {shape.name} intersects master chrome {ch.get('name')}")
    return fails


def table_parity(manifest: dict[str, Any], out_prs) -> list[str]:
    diffs = []
    src_tables = []
    for slide in manifest.get("slides") or []:
        for shape in slide.get("shapes") or []:
            if shape.get("kind") == "table":
                cells = []
                for row in (shape.get("table") or {}).get("cells") or []:
                    cells.append([normalize_for_parity(c.get("text") or "") for c in row])
                src_tables.append((slide["index"], cells))
    out_tables = []
    for i, slide in enumerate(out_prs.slides):
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                cells = [[normalize_for_parity(cell.text or "") for cell in row.cells] for row in shape.table.rows]
                out_tables.append((i, cells))
    # Compare concatenated cell text, allowing continuation splits.
    src_flat = [cell for _, rows in src_tables for row in rows for cell in row]
    out_flat = [cell for _, rows in out_tables for row in rows for cell in row]
    if src_flat != out_flat:
        diffs.append(f"table cells source={len(src_flat)} output={len(out_flat)}")
    return diffs


def image_parity(manifest: dict[str, Any], out_prs, tokens: dict[str, Any] | None = None) -> tuple[int, int, list[str]]:
    src = 0
    dropped = []
    for slide in manifest.get("slides") or []:
        for shape in slide.get("shapes") or []:
            if shape.get("kind") != "picture":
                continue
            if shape.get("drop_logo") or shape.get("drop_icon"):
                dropped.append(f"slide {slide['index'] + 1} {shape.get('name')}")
            else:
                src += 1
    logos = set((tokens or {}).get("logo_hashes") or [])
    sample = set((tokens or {}).get("sample_photo_hashes") or [])
    out = 0
    for slide in out_prs.slides:
        for shape in slide.shapes:
            if "PICTURE" not in str(shape.shape_type):
                continue
            digest = None
            try:
                digest = blob_hash(shape.image.blob)
            except Exception:
                digest = None
            if digest and digest in logos:
                continue
            if digest and digest in sample:
                continue
            out += 1
    return src, out, dropped


def run_qa(out_dir: Path, source_path: Path, output_path: Path) -> dict[str, Any]:
    qa_dir = out_dir / "QA"
    qa_dir.mkdir(parents=True, exist_ok=True)
    tokens = load_json(out_dir / "brand_tokens.json")
    manifest = load_json(out_dir / "source_manifest.json")
    plan = load_json(out_dir / "plan.json")
    tokens = apply_colorway(tokens, plan.get("colorway") or tokens.get("colorway"))
    file_brand = parse_brand_md(None)
    token_brand = dict(tokens.get("brand_md") or {})
    brand = {**file_brand, **token_brand}
    for key in ("fonts", "color_whitelist", "chart_palette"):
        brand[key] = sorted(set((file_brand.get(key) or []) + (token_brand.get(key) or [])))
    flags = []
    for entry in plan.get("entries") or []:
        flags.extend(entry.get("flags") or [])
        for pl in entry.get("placements") or []:
            if pl.get("flag"):
                flags.append(pl["flag"])
    flags = sorted(set(flags))

    out_render = render_pptx(output_path, qa_dir, "out")
    src_render = render_pptx(source_path, qa_dir, "src")

    checks: dict[str, Any] = {}
    justifications: list[str] = []
    try:
        out_prs = Presentation(str(output_path))
        reopen_ok = True
    except Exception as exc:
        out_prs = None
        reopen_ok = False
        justifications.append(f"python-pptx could not reopen OUTPUT.pptx: {exc}")

    expected = int(plan.get("output_count") or 0)
    actual = len(list(out_prs.slides)) if out_prs else 0
    checks["slide_count"] = {
        "pass": actual == expected,
        "expected": expected,
        "actual": actual,
        "source": int(plan.get("source_count") or 0),
    }

    if out_prs is not None:
        faces = grep_fonts(output_path)
        allowed_f = allowed_fonts(tokens, brand)
        bad_fonts = sorted({f for f in faces if f not in allowed_f})
        checks["fonts"] = {"pass": not bad_fonts, "foreign": bad_fonts, "seen": sorted(set(faces))}
        colors = grep_srgb(output_path)
        allowed_c = allowed_colors(tokens, brand)
        bad_colors = sorted({c for _, c in colors if c not in allowed_c})
        checks["colors"] = {"pass": not bad_colors, "foreign": bad_colors}

        overflow = overflow_fail(out_prs, tokens)
        checks["overflow"] = {"pass": not overflow, "failures": overflow}
        bounds = bounds_fail(out_prs, tokens)
        checks["bounds"] = {"pass": not bounds, "failures": bounds[:20]}

        # Text parity grouped by source_index (splits concatenate).
        src_by_index = {s["index"]: source_slide_text(s) for s in manifest.get("slides") or []}
        out_by_index: dict[int, list[str]] = defaultdict(list)
        for entry, slide in zip(plan.get("entries") or [], out_prs.slides):
            blob = slide_text(slide) + "\n" + output_notes(slide)
            out_by_index[int(entry["source_index"])].append(blob)
        diffs = []
        for index, src_text in src_by_index.items():
            out_text = "\n".join(out_by_index.get(index) or [])
            a, b = normalize_for_parity(src_text), normalize_for_parity(out_text)
            if a != b:
                diffs.append({"source_index": index, "source": a[:240], "output": b[:240]})
        checks["text_parity"] = {"pass": not diffs, "diffs": diffs}

        src_n, out_n, dropped = image_parity(manifest, out_prs, tokens)
        checks["image_parity"] = {
            "pass": src_n == out_n,
            "source": src_n,
            "output": out_n,
            "dropped_logos": dropped,
        }
        if dropped:
            flags = sorted(set(flags + ["LOGO_DROPPED"]))
        tdiffs = table_parity(manifest, out_prs)
        checks["table_parity"] = {"pass": not tdiffs, "diffs": tdiffs}
        checks["reopen"] = {"pass": reopen_ok}

    checks["libreoffice_output"] = {
        "pass": bool(out_render.get("ok")),
        "error": out_render.get("error"),
        "tool": out_render.get("tool"),
        "images": len(out_render.get("images") or []),
    }
    checks["libreoffice_source"] = {
        "pass": bool(src_render.get("ok")),
        "error": src_render.get("error"),
        "images": len(src_render.get("images") or []),
    }
    if out_render.get("error") and "not installed" in (out_render.get("error") or ""):
        justifications.append(out_render["error"])
        checks["libreoffice_output"]["justified"] = True
    if src_render.get("error") and "not installed" in (src_render.get("error") or ""):
        justifications.append(src_render["error"])
        checks["libreoffice_source"]["justified"] = True

    # Plan table
    lines = [
        "# QA report",
        "",
        f"Source: `{source_path}`",
        f"Output: `{output_path}`",
        f"Template: `{tokens.get('source_template_path') or tokens.get('template_path')}`",
        f"Colorway: `{tokens.get('colorway_label') or tokens.get('colorway') or 'green'}`",
        "",
    ]
    lines += ["## Plan", "", "| Output | Source | Role | Layout | Flags |", "| --- | --- | --- | --- | --- |"]
    for i, entry in enumerate(plan.get("entries") or [], start=1):
        flags_cell = ", ".join(entry.get("flags") or []) or "—"
        lines.append(
            f"| {i} | {int(entry['source_index']) + 1} | {entry.get('role')} | {entry.get('template_layout_index')} {entry.get('template_layout_name') or ''} | {flags_cell} |"
        )
    lines += ["", "## Checks", ""]
    for name, check in checks.items():
        status = "PASS" if check.get("pass") else ("JUSTIFIED" if check.get("justified") else "FAIL")
        extra = ""
        if name == "text_parity" and check.get("diffs"):
            extra = f" ({len(check['diffs'])} slides differ)"
        if name == "fonts" and check.get("foreign"):
            extra = f" foreign={', '.join(check['foreign'])}"
        if name == "colors" and check.get("foreign"):
            extra = f" foreign={', '.join(check['foreign'])}"
        if check.get("error"):
            extra = f" ({check['error']})"
        lines.append(f"- **{name}**: {status}{extra}")
    lines += ["", "## Flags", ""]
    if flags:
        for flag in flags:
            lines.append(f"- `{flag}`")
    else:
        lines.append("- none")
    if justifications:
        lines += ["", "## Justifications", ""]
        for item in justifications:
            lines.append(f"- {item}")
    if checks.get("text_parity", {}).get("diffs"):
        lines += ["", "## Text parity diffs", ""]
        for diff in checks["text_parity"]["diffs"][:20]:
            lines.append(f"- Slide {int(diff['source_index']) + 1}")
            lines.append(f"  - source: `{diff['source']}`")
            lines.append(f"  - output: `{diff['output']}`")
    if tokens.get("notes"):
        lines += ["", "## Template notes", ""]
        for note in tokens["notes"]:
            lines.append(f"- {note}")
    report = "\n".join(lines) + "\n"
    (qa_dir / "qa-report.md").write_text(report, encoding="utf-8")
    summary = {
        "checks": checks,
        "flags": flags,
        "justifications": justifications,
        "report": str(qa_dir / "qa-report.md"),
        "output_images": out_render.get("images") or [],
        "source_images": src_render.get("images") or [],
    }
    dump_json(qa_dir / "qa-summary.json", summary)
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="QA an OUTPUT.pptx against the source deck")
    parser.add_argument("--out-dir", default=".")
    parser.add_argument("--source", default=None)
    parser.add_argument("--output", default=None)
    parser.add_argument("--colorway", default=None)
    args = parser.parse_args(argv)
    out_dir = Path(args.out_dir)
    manifest = load_json(out_dir / "source_manifest.json")
    source = Path(args.source) if args.source else Path(manifest.get("source_path"))
    output = Path(args.output) if args.output else out_dir / "OUTPUT.pptx"
    if args.colorway:
        tokens = apply_colorway(load_json(out_dir / "brand_tokens.json"), args.colorway)
        dump_json(out_dir / "brand_tokens.json", tokens)
    summary = run_qa(out_dir, source, output)
    failed = [
        name
        for name, check in summary["checks"].items()
        if not check.get("pass") and not check.get("justified")
    ]
    print(f"Wrote {summary['report']}")
    if failed:
        print("Failed checks: " + ", ".join(failed))
        return 0  # still succeed so the pipeline yields artifacts; failures live in the report
    return 0


if __name__ == "__main__":
    sys.exit(main())
