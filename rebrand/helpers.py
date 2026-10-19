"""Spec §6 helpers plus typography, geometry, table, chart, and image utilities.

Build never XML-deepcopies charts and never paints per-slide brand chrome.
"""

from __future__ import annotations

import hashlib
import io
import json
import math
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any, Iterable
from xml.sax.saxutils import escape

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

NSMAP = {"a": A_NS, "p": P_NS, "r": R_NS, "c": C_NS, "dgm": DGM_NS, "p14": P14_NS}

ROLES = (
    "TITLE",
    "SECTION",
    "TITLE_BODY",
    "TWO_CONTENT",
    "TITLE_ONLY",
    "PICTURE",
    "BLANK",
    "CLOSING",
)

ROLE_LABELS = {
    "TITLE": "Title",
    "SECTION": "Section",
    "TITLE_BODY": "Title and body",
    "TWO_CONTENT": "Two content",
    "TITLE_ONLY": "Title only",
    "PICTURE": "Picture",
    "BLANK": "Blank",
    "CLOSING": "Closing",
}

ROLE_FALLBACKS = {
    "TWO_CONTENT": ["TITLE_BODY", "TITLE_ONLY", "BLANK"],
    "PICTURE": ["TITLE_ONLY", "BLANK"],
    "SECTION": ["TITLE", "TITLE_ONLY", "BLANK"],
    "CLOSING": ["TITLE", "TITLE_ONLY", "BLANK"],
    "TITLE_BODY": ["TITLE_ONLY", "BLANK"],
    "TITLE": ["TITLE_ONLY", "BLANK"],
    "TITLE_ONLY": ["BLANK"],
    "BLANK": ["TITLE_ONLY"],
}

TITLE_PH = {
    PP_PLACEHOLDER.TITLE,
    PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.VERTICAL_TITLE,
}
BODY_PH = {
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.VERTICAL_BODY,
}
SUBTITLE_PH = {PP_PLACEHOLDER.SUBTITLE}
PICTURE_PH = {PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.MEDIA_CLIP}
FOOTER_PH = {
    PP_PLACEHOLDER.SLIDE_NUMBER,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.DATE,
}

FOOTER_TEXT_RE = re.compile(
    r"(^\d{1,3}$)|(^(page|slide)\s*\d+)|©|confidential",
    re.IGNORECASE,
)
MANUAL_BULLET_RE = re.compile(r"^[\s]*([•●▪◦○■□–—\-]|o )[\s]+")
MANUAL_NUMBER_RE = re.compile(r"^[\s]*\d+[.)]\s+")
CONTACT_RE = re.compile(
    r"\b(thank you|thanks|questions|q\s*&\s*a|contact|goodbye|the end)\b",
    re.IGNORECASE,
)

POTX_MAIN = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
PPTX_MAIN = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"

EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def hex_color(value: str | None, default: str = "000000") -> str:
    raw = (value or default).strip().lstrip("#").upper()
    if len(raw) == 3:
        raw = "".join(ch * 2 for ch in raw)
    if len(raw) != 6 or any(ch not in "0123456789ABCDEF" for ch in raw):
        return default.upper()
    return raw


def rgb(hex_value: str) -> RGBColor:
    h = hex_color(hex_value)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def tint(hex_color_value: str, pct: float) -> str:
    """pct 0..1 toward white."""
    h = hex_color(hex_color_value)
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    return "".join(f"{int(c + (255 - c) * pct):02X}" for c in (r, g, b))


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def dump_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def parse_brand_md(path: Path | None) -> dict[str, Any]:
    """Lightweight brand.md reader. Unknown keys are kept as strings."""
    rules: dict[str, Any] = {"fonts": [], "chart_palette": [], "color_whitelist": []}
    if path is None or not path.exists():
        return rules
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip().lower().replace(" ", "_")
        value = value.strip().strip("`").strip('"').strip("'")
        if key in {"chart_palette", "palette", "color_whitelist", "fonts"}:
            items = [hex_color(part) if key != "fonts" else part.strip() for part in re.split(r"[,\[\]\s]+", value) if part.strip()]
            rules[key if key != "palette" else "chart_palette"] = items
        elif key in {"east_asian_font", "ea_font", "complex_font", "cs_font"}:
            dest = "east_asian_font" if "ea" in key or "east" in key else "complex_font"
            rules[dest] = value
        elif key in {"rounded_corners", "image_rounded_corners"}:
            rules["rounded_corners"] = value.lower() in {"1", "true", "yes", "on"}
        elif key in {"image_border", "picture_border"}:
            rules["image_border"] = value.lower() in {"1", "true", "yes", "on"}
        else:
            rules[key] = value
    return rules


def ensure_pptx(path: Path, dest: Path | None = None) -> Path:
    """Copy .potx → .pptx and rewrite the main part content type."""
    path = Path(path)
    if path.suffix.lower() == ".pptx":
        return path
    dest = Path(dest) if dest is not None else path.with_suffix(".pptx")
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, dest)
    buf = io.BytesIO(dest.read_bytes())
    with zipfile.ZipFile(buf, "r") as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}
    ct = contents.get("[Content_Types].xml", b"")
    ct = ct.replace(POTX_MAIN.encode("utf-8"), PPTX_MAIN.encode("utf-8"))
    contents["[Content_Types].xml"] = ct
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in contents.items():
            zout.writestr(name, data)
    dest.write_bytes(out.getvalue())
    return dest


def delete_slide(prs, index: int) -> None:
    sld_id_lst = prs.slides._sldIdLst
    sld_id = sld_id_lst[index]
    prs.part.drop_rel(sld_id.rId)
    sld_id_lst.remove(sld_id)


def delete_all_slides(prs) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        prs.part.drop_rel(sld_id.rId)
        sld_id_lst.remove(sld_id)


def estimate_lines(text: str, font_pt: float, box_width_emu: int, avg_char_w: float = 0.5) -> int:
    chars_per_line = max(1, int((box_width_emu / EMU_PER_PT) / (font_pt * avg_char_w)))
    return sum(max(1, -(-len(p) // chars_per_line)) for p in (text or "").split("\n"))


def fits(paragraphs: Iterable[Any], font_pt: float, box_w: int, box_h: int, line_spacing: float = 1.2) -> bool:
    texts: list[str] = []
    for item in paragraphs:
        if isinstance(item, str):
            texts.append(item)
        elif isinstance(item, dict):
            texts.append(item.get("text") or "")
        else:
            texts.append(getattr(item, "text", "") or "")
    lines = sum(estimate_lines(text, font_pt, box_w) for text in texts) or 1
    return lines * font_pt * line_spacing * EMU_PER_PT <= box_h


def set_cell_border(cell, side: str, width_pt: float, hex_color_value: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tag = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}[side]
    for elem in tc_pr.findall(qn(tag)):
        tc_pr.remove(elem)
    ln = etree.SubElement(
        tc_pr,
        qn(tag),
        w=str(int(width_pt * EMU_PER_PT)),
        cap="flat",
        cmpd="sng",
        algn="ctr",
    )
    sf = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val=hex_color(hex_color_value))
    etree.SubElement(ln, qn("a:prstDash"), val="solid")


def clear_cell_border(cell, side: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tag = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}[side]
    for elem in tc_pr.findall(qn(tag)):
        tc_pr.remove(elem)
    ln = etree.SubElement(tc_pr, qn(tag), w="0")
    etree.SubElement(ln, qn("a:noFill"))


def clear_table_style(table) -> None:
    tbl_pr = table._tbl.tblPr
    for elem in tbl_pr.findall(qn("a:tableStyleId")):
        tbl_pr.remove(elem)
    tbl_pr.set("firstRow", "1")
    tbl_pr.set("bandRow", "0")
    tbl_pr.set("firstCol", "0")
    tbl_pr.set("lastRow", "0")
    tbl_pr.set("lastCol", "0")
    tbl_pr.set("bandCol", "0")


def strip_run(run) -> None:
    r_pr = run._r.get_or_add_rPr()
    for attr in ("sz", "b", "i", "u", "strike", "baseline", "cap", "spc"):
        r_pr.attrib.pop(attr, None)
    for child in list(r_pr):
        local = child.tag.split("}")[-1]
        if local not in {"hlinkClick"}:
            r_pr.remove(child)


def set_run_typeface(run, latin: str, east_asian: str | None = None, complex_script: str | None = None) -> None:
    r_pr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", latin), ("a:ea", east_asian or latin), ("a:cs", complex_script or latin)):
        node = r_pr.find(qn(tag))
        if node is None:
            node = etree.SubElement(r_pr, qn(tag))
        node.set("typeface", face)


def placeholder_type(shape) -> Any | None:
    try:
        return shape.placeholder_format.type
    except Exception:
        return None


def placeholder_type_name(shape) -> str | None:
    ph = placeholder_type(shape)
    return enum_name(ph) or None


def placeholder_idx(shape) -> int | None:
    try:
        return int(shape.placeholder_format.idx)
    except Exception:
        return None


def iter_shapes(shapes) -> list[Any]:
    found = []
    for shape in shapes:
        found.append(shape)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            found.extend(iter_shapes(shape.shapes))
    return found


def iter_shapes_abs(shapes, dx: int = 0, dy: int = 0) -> list[tuple[Any, int, int]]:
    found: list[tuple[Any, int, int]] = []
    for shape in shapes:
        left = int(dx + int(getattr(shape, "left", 0) or 0))
        top = int(dy + int(getattr(shape, "top", 0) or 0))
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            found.append((shape, left, top))
            found.extend(iter_shapes_abs(shape.shapes, left, top))
        else:
            found.append((shape, left, top))
    return found


def shape_kind(shape) -> str:
    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if st in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
        return "picture"
    if st in {MSO_SHAPE_TYPE.DIAGRAM, MSO_SHAPE_TYPE.IGX_GRAPHIC}:
        return "smartart"
    ph = placeholder_type(shape)
    if ph == PP_PLACEHOLDER.PICTURE:
        return "picture"
    if st in {MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT}:
        if _graphic_uri(shape).endswith("diagram"):
            return "smartart"
        if getattr(shape, "has_chart", False):
            return "chart"
        return "video" if st == MSO_SHAPE_TYPE.MEDIA else "other"
    if getattr(shape, "has_chart", False) or st == MSO_SHAPE_TYPE.CHART:
        return "chart"
    if getattr(shape, "has_table", False):
        return "table"
    uri = _graphic_uri(shape)
    if "diagram" in uri:
        return "smartart"
    if st in {MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.FREEFORM}:
        return "other"
    if getattr(shape, "has_text_frame", False):
        return "text"
    return "other"


def _graphic_uri(shape) -> str:
    try:
        node = shape._element.find(f".//{{{A_NS}}}graphicData")
        return (node.get("uri") or "") if node is not None else ""
    except Exception:
        return ""


def is_line_or_arrow(shape) -> bool:
    if shape.shape_type == MSO_SHAPE_TYPE.LINE:
        return True
    name = (getattr(shape, "name", "") or "").lower()
    if any(token in name for token in ("arrow", "connector", "line")):
        return True
    try:
        prst = shape.auto_shape_type
    except Exception:
        return False
    label = str(prst).lower()
    return "arrow" in label or "line" in label or "connector" in label


def chart_is_external(shape) -> bool:
    try:
        xlsx = shape.chart.part.chart_workbook.xlsx_part
    except Exception:
        return True
    return xlsx is None


def theme_color(tokens: dict[str, Any], name: str, default: str = "000000") -> str:
    colors = tokens.get("theme", {}).get("colors") or tokens.get("colors") or {}
    return hex_color(colors.get(name), default)


def theme_font(tokens: dict[str, Any], which: str = "minor") -> str:
    fonts = tokens.get("theme", {}).get("fonts") or tokens.get("fonts") or {}
    if which == "major":
        return fonts.get("major") or fonts.get("minor") or "Calibri"
    return fonts.get("minor") or fonts.get("major") or "Calibri"


def body_size_for_level(level: int) -> int:
    if level <= 0:
        return 18
    if level == 1:
        return 16
    return 14


def reduce_body_size(paragraphs: list[dict[str, Any]], box_w: int, box_h: int) -> int:
    for size in (18, 16, 14, 12):
        scaled = []
        for para in paragraphs:
            level = int(para.get("level") or 0)
            pt = body_size_for_level(level)
            use = min(pt, size) if level == 0 else min(pt, max(12, size - level * 2))
            scaled.append({"text": para.get("text") or ""})
            _ = use
        # Estimate using the level-0 size as the dominant metric.
        if fits(paragraphs, size, box_w, box_h):
            return size
    return 12


def reduce_title_size(text: str, box_w: int, box_h: int, start: int) -> int:
    size = start
    while size > 24 and not fits([text], size, box_w, box_h, line_spacing=1.1):
        size -= 2
    return max(24, size)


def normalize_alignment(value: Any) -> str:
    label = str(value or "").lower()
    if "center" in label:
        return "center"
    if "right" in label:
        return "left"
    if "justify" in label:
        return "left"
    return "left"


def pp_align(value: str):
    return PP_ALIGN.CENTER if value == "center" else PP_ALIGN.LEFT


def looks_number(text: str) -> bool:
    cleaned = text.strip().replace(",", "").replace("$", "").replace("%", "")
    if not cleaned:
        return False
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def clean_paragraph_text(text: str, numbered: bool = False) -> tuple[str, bool]:
    """Strip manual bullet/number glyphs. Returns (text, forced_bullet)."""
    raw = text or ""
    forced = False
    if MANUAL_BULLET_RE.match(raw):
        raw = MANUAL_BULLET_RE.sub("", raw, count=1)
        forced = True
    if numbered or MANUAL_NUMBER_RE.match(raw):
        raw = MANUAL_NUMBER_RE.sub("", raw, count=1)
        forced = forced or numbered
    return raw, forced


def apply_quote_style(text: str, curly: bool) -> str:
    if not curly:
        return text
    out = []
    opening = True
    for ch in text:
        if ch == '"':
            out.append("“" if opening else "”")
            opening = not opening
        elif ch == "'":
            out.append("‘" if opening else "’")
            opening = not opening
        else:
            out.append(ch)
    return "".join(out)


def apply_run_style(
    run,
    tokens: dict[str, Any],
    brand: dict[str, Any],
    *,
    font_kind: str,
    size_pt: float,
    color_hex: str,
    bold: bool | None,
    italic: bool | None,
    hyperlink: str | None = None,
    underline_if_link: bool = True,
) -> None:
    name = theme_font(tokens, "major" if font_kind == "major" else "minor")
    ea = brand.get("east_asian_font")
    cs = brand.get("complex_font")
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    run.font.color.rgb = rgb(color_hex)
    run.font.name = name
    set_run_typeface(run, name, ea, cs)
    if hyperlink:
        try:
            run.hyperlink.address = hyperlink
        except Exception:
            pass
        run.font.color.rgb = rgb(theme_color(tokens, "hlink", color_hex))
        if underline_if_link:
            run.font.underline = True


def write_text_frame(
    tf,
    paragraphs: list[dict[str, Any]],
    tokens: dict[str, Any],
    brand: dict[str, Any],
    *,
    kind: str,
    role: str = "TITLE_BODY",
    body_size: int | None = None,
    title_size: int | None = None,
    curly_quotes: bool = False,
) -> None:
    tf.clear()
    tf.word_wrap = True
    if not paragraphs:
        return
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        level = int(para.get("level") or 0)
        p.level = max(0, min(level, 8))
        p.alignment = pp_align(normalize_alignment(para.get("alignment")))
        numbered = bool(para.get("numbered"))
        runs = para.get("runs") or [{"text": para.get("text") or "", "bold": para.get("bold"), "italic": para.get("italic")}]
        if kind in {"body", "subtitle", "caption", "footnote", "callout"}:
            size = body_size or body_size_for_level(level)
            if kind == "subtitle":
                font_kind, size, color, default_bold, default_italic = "minor", 20, theme_color(tokens, "dk2"), False, False
            elif kind == "caption":
                font_kind, size, color, default_bold, default_italic = "minor", 10, theme_color(tokens, "dk2"), False, True
            elif kind == "footnote":
                font_kind, size, color, default_bold, default_italic = "minor", 9, theme_color(tokens, "dk2"), False, False
            elif kind == "callout":
                font_kind, size, color, default_bold, default_italic = "minor", body_size_for_level(level), theme_color(tokens, "dk1"), False, False
            else:
                font_kind = "minor"
                color = theme_color(tokens, "dk1")
                default_bold, default_italic = False, False
                if level == 0:
                    size = body_size or 18
                elif level == 1:
                    size = min(16, body_size or 16)
                else:
                    size = min(14, body_size or 14)
        elif kind == "textbox":
            font_kind = "minor"
            color = theme_color(tokens, "dk1")
            default_bold, default_italic = False, False
            size = body_size or body_size_for_level(level)
        elif kind == "cover_title":
            sample = (tokens.get("sample_title_slide_fonts") or {}).get("title") or {}
            font_kind = "major"
            size = title_size or int(sample.get("size_pt") or 40)
            color = hex_color(sample.get("color"), theme_color(tokens, "dk1"))
            default_bold = bool(sample.get("bold", True))
            default_italic = False
        elif kind == "section_title":
            font_kind, size, color, default_bold, default_italic = "major", title_size or 36, theme_color(tokens, "dk1"), True, False
        else:
            sample = ((tokens.get("layout_title_fonts") or {}).get(str(tokens.get("_current_layout"))) or {})
            font_kind = "major"
            size = title_size or int(sample.get("size_pt") or 28)
            color = theme_color(tokens, "dk1")
            default_bold = True if sample.get("bold") is None else bool(sample.get("bold"))
            default_italic = False

        wrote = False
        for run_info in runs:
            text = apply_quote_style(run_info.get("text") or "", curly_quotes)
            text, _forced = clean_paragraph_text(text, numbered=numbered)
            if not text and wrote:
                continue
            run = p.add_run()
            run.text = text
            link = run_info.get("hyperlink")
            bold = run_info.get("bold")
            italic = run_info.get("italic")
            if bold is None:
                bold = default_bold
            if italic is None:
                italic = default_italic
            apply_run_style(
                run,
                tokens,
                brand,
                font_kind=font_kind,
                size_pt=size,
                color_hex=theme_color(tokens, "hlink") if link else color,
                bold=bool(bold) if bold is not None else default_bold,
                italic=bool(italic) if italic is not None else default_italic,
                hyperlink=link,
            )
            wrote = True
        if not wrote:
            run = p.add_run()
            run.text = apply_quote_style(para.get("text") or "", curly_quotes)
            apply_run_style(
                run,
                tokens,
                brand,
                font_kind=font_kind,
                size_pt=size,
                color_hex=color,
                bold=default_bold,
                italic=default_italic,
            )


def set_cell_margins(cell, left_in=0.05, right_in=0.05, top_in=0.03, bottom_in=0.03) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_pr.set("marL", str(int(Inches(left_in))))
    tc_pr.set("marR", str(int(Inches(right_in))))
    tc_pr.set("marT", str(int(Inches(top_in))))
    tc_pr.set("marB", str(int(Inches(bottom_in))))


def fill_cell(cell, hex_value: str) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb(hex_value)


def as_box(left: int, top: int, width: int, height: int) -> dict[str, int]:
    return {"left": int(left), "top": int(top), "width": max(1, int(width)), "height": max(1, int(height))}


def box_right(box: dict[str, int]) -> int:
    return int(box["left"]) + int(box["width"])


def box_bottom(box: dict[str, int]) -> int:
    return int(box["top"]) + int(box["height"])


def intersect_box(a: dict[str, int], b: dict[str, int]) -> dict[str, int] | None:
    left = max(int(a["left"]), int(b["left"]))
    top = max(int(a["top"]), int(b["top"]))
    right = min(box_right(a), box_right(b))
    bottom = min(box_bottom(a), box_bottom(b))
    if right <= left or bottom <= top:
        return None
    return as_box(left, top, right - left, bottom - top)


def overlap_fraction(a: dict[str, int], b: dict[str, int]) -> float:
    inter = intersect_box(a, b)
    if inter is None:
        return 0.0
    area_a = max(1, int(a["width"]) * int(a["height"]))
    return (int(inter["width"]) * int(inter["height"])) / area_a


def clamp_box_to_slide(box: dict[str, int], slide_w: int, slide_h: int) -> dict[str, int]:
    slide = as_box(0, 0, int(slide_w), int(slide_h))
    clipped = intersect_box(box, slide)
    return clipped or as_box(0, 0, 1, 1)


def is_brand_lockup(
    area: int,
    slide_area: int,
    left: int,
    top: int,
    width: int,
    height: int,
    slide_w: int,
    slide_h: int,
) -> bool:
    """Small top-left (or edge) brand mark — not a content photo frame."""
    if area > slide_area * 0.05:
        return False
    if left <= slide_w * 0.25 and top <= slide_h * 0.30:
        return True
    return is_edge_logo(area, slide_area, left, top, width, height, slide_w, slide_h)


def _prototype_record(tokens: dict[str, Any], layout_index: int | None) -> dict[str, Any] | None:
    if layout_index is None:
        return None
    for proto in tokens.get("prototypes") or []:
        try:
            if int(proto.get("index")) == int(layout_index):
                return proto
        except (TypeError, ValueError):
            continue
    return None


def content_safe_box(tokens: dict[str, Any], layout_index: int | None = None) -> dict[str, int]:
    width = int(tokens["slide_width"])
    height = int(tokens["slide_height"])
    proto = _prototype_record(tokens, layout_index)
    if proto and proto.get("content_box"):
        box = {k: int(proto["content_box"][k]) for k in ("left", "top", "width", "height")}
        return _shrink_for_chrome(clamp_box_to_slide(box, width, height), tokens)
    inset = tokens.get("content_inset") or {}
    side = int(inset.get("left") or width * 0.05)
    top = int(inset.get("top") or height * 0.16)
    bottom_margin = int(inset.get("bottom") or height * 0.08)
    right_margin = int(inset.get("right") or side)
    layouts = tokens.get("layouts") or []
    layout = None
    if layout_index is not None and 0 <= layout_index < len(layouts):
        layout = layouts[layout_index]
    if layout:
        bodies = [
            ph
            for ph in layout.get("placeholders") or []
            if ph.get("type") in {str(t) for t in BODY_PH} or (ph.get("type_name") or "").upper() in {"BODY", "OBJECT", "VERTICAL_BODY"}
        ]
        if not bodies:
            bodies = [
                ph
                for ph in layout.get("placeholders") or []
                if (ph.get("type_name") or "").upper() in {"BODY", "OBJECT", "SUBTITLE"}
            ]
        if bodies:
            lefts = [int(ph["left"]) for ph in bodies]
            tops = [int(ph["top"]) for ph in bodies]
            rights = [int(ph["left"]) + int(ph["width"]) for ph in bodies]
            bottoms = [int(ph["top"]) + int(ph["height"]) for ph in bodies]
            box = {
                "left": min(lefts),
                "top": min(tops),
                "width": max(rights) - min(lefts),
                "height": max(bottoms) - min(tops),
            }
            return _shrink_for_chrome(clamp_box_to_slide(box, width, height), tokens)
    box = {"left": side, "top": top, "width": width - side - right_margin, "height": height - top - bottom_margin}
    return _shrink_for_chrome(clamp_box_to_slide(box, width, height), tokens)


def data_content_box(tokens: dict[str, Any], layout_index: int | None = None) -> dict[str, int]:
    """Full content well under the title band, down to the slide floor.

    Comparison sample art only paints five dummy rows. When a table (or chart)
    *is* the slide, use the remaining canvas instead of that leftover strip.
    """
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    safe = content_safe_box(tokens, layout_index)
    top = max(int(safe["top"]), int(slide_h * 0.22))
    floor = slide_h - int(0.32 * EMU_PER_INCH)
    return clamp_box_to_slide(
        as_box(int(safe["left"]), top, int(safe["width"]), max(int(1.2 * EMU_PER_INCH), floor - top)),
        slide_w,
        slide_h,
    )


def _shrink_for_chrome(box: dict[str, int], tokens: dict[str, Any]) -> dict[str, int]:
    left, top, width, height = box["left"], box["top"], box["width"], box["height"]
    right, bottom = left + width, top + height
    for chrome in tokens.get("master_chrome") or []:
        c_left, c_top = int(chrome["left"]), int(chrome["top"])
        c_right = c_left + int(chrome["width"])
        c_bottom = c_top + int(chrome["height"])
        # If chrome sits along the bottom, raise the content floor.
        slide_h = int(tokens["slide_height"])
        if c_top > slide_h * 0.85:
            bottom = min(bottom, c_top)
        elif c_bottom < slide_h * 0.15:
            top = max(top, c_bottom)
        elif c_right < int(tokens["slide_width"]) * 0.15:
            left = max(left, c_right)
        elif c_left > int(tokens["slide_width"]) * 0.85:
            right = min(right, c_left)
    width = max(1, right - left)
    height = max(1, bottom - top)
    return {"left": int(left), "top": int(top), "width": int(width), "height": int(height)}


def scale_geometry(src_box: dict[str, int], src_size: dict[str, int], tokens: dict[str, Any], layout_index: int | None) -> dict[str, int]:
    tpl_w = int(tokens["slide_width"])
    tpl_h = int(tokens["slide_height"])
    src_w = max(1, int(src_size.get("width") or tpl_w))
    src_h = max(1, int(src_size.get("height") or tpl_h))
    s = min(tpl_w / src_w, tpl_h / src_h)
    safe = content_safe_box(tokens, layout_index)
    left = int(round(int(src_box.get("left") or 0) * s))
    top = int(round(int(src_box.get("top") or 0) * s))
    width = int(round(max(1, int(src_box.get("width") or 1)) * s))
    height = int(round(max(1, int(src_box.get("height") or 1)) * s))
    # Center the scaled block horizontally in the content-safe area and snap to body top.
    block_left = int(src_box.get("left") or 0)
    # Keep relative placement but clamp into safe area.
    if left < safe["left"]:
        left = safe["left"]
    if left + width > safe["left"] + safe["width"]:
        left = max(safe["left"], safe["left"] + safe["width"] - width)
        if width > safe["width"]:
            height = int(height * (safe["width"] / width))
            width = safe["width"]
            left = safe["left"]
    if top < safe["top"]:
        top = safe["top"]
    if top + height > safe["top"] + safe["height"]:
        top = max(safe["top"], safe["top"] + safe["height"] - height)
        if height > safe["height"]:
            width = int(width * (safe["height"] / height))
            height = safe["height"]
            top = safe["top"]
    return {"left": int(left), "top": int(top), "width": max(1, int(width)), "height": max(1, int(height))}


def contain_fit(img_w: int, img_h: int, box: dict[str, int]) -> dict[str, int]:
    img_w = max(1, img_w)
    img_h = max(1, img_h)
    s = min(box["width"] / img_w, box["height"] / img_h)
    width = max(1, int(img_w * s))
    height = max(1, int(img_h * s))
    left = int(box["left"] + (box["width"] - width) / 2)
    top = int(box["top"] + (box["height"] - height) / 2)
    return {"left": left, "top": top, "width": width, "height": height}


def blob_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def maybe_recompress(blob: bytes, ext: str) -> tuple[bytes, str]:
    if len(blob) <= 4 * 1024 * 1024:
        return blob, ext
    try:
        from PIL import Image
    except Exception:
        return blob, ext
    im = Image.open(io.BytesIO(blob))
    out = io.BytesIO()
    if im.mode in {"RGBA", "LA"} or "A" in im.getbands():
        im.save(out, format="PNG", optimize=True)
        return out.getvalue(), "png"
    rgb_im = im.convert("RGB")
    rgb_im.save(out, format="JPEG", quality=85, optimize=True)
    return out.getvalue(), "jpg"


def picture_crop(shape) -> dict[str, int]:
    crop = {"left": 0, "right": 0, "top": 0, "bottom": 0}
    try:
        src = shape._element.find(f".//{{{A_NS}}}srcRect")
        if src is None:
            return crop
        mapping = {"l": "left", "r": "right", "t": "top", "b": "bottom"}
        for attr, key in mapping.items():
            crop[key] = int(src.get(attr, "0") or 0)
    except Exception:
        return crop
    return crop


def apply_picture_crop(shape, crop: dict[str, int]) -> None:
    if not any(crop.get(k) for k in ("left", "right", "top", "bottom")):
        return
    blip_fill = shape._element.find(f".//{{{A_NS}}}blipFill")
    if blip_fill is None:
        return
    src = blip_fill.find(qn("a:srcRect"))
    if src is None:
        src = etree.SubElement(blip_fill, qn("a:srcRect"))
    src.set("l", str(int(crop.get("left") or 0)))
    src.set("r", str(int(crop.get("right") or 0)))
    src.set("t", str(int(crop.get("top") or 0)))
    src.set("b", str(int(crop.get("bottom") or 0)))


def crop_cut_fraction(crop: dict[str, int]) -> float:
    # srcRect values are 1000ths of a percent (so 100000 = 100%).
    left = (crop.get("left") or 0) / 100000
    right = (crop.get("right") or 0) / 100000
    top = (crop.get("top") or 0) / 100000
    bottom = (crop.get("bottom") or 0) / 100000
    remain = max(0.0, 1 - left - right) * max(0.0, 1 - top - bottom)
    return max(0.0, 1 - remain)


ORG_CHROME_RE = re.compile(
    r"^(the\s+)?everglades\s+foundation(\s*[·•|\-—:]\s+\S+(?:\s+\S+){0,6})?$",
    re.I,
)


def is_org_chrome(text: str) -> bool:
    """True for org lockup lines, including 'ORG · PROGRAM' kickers."""
    t = " ".join((text or "").split())
    return bool(ORG_CHROME_RE.match(t))


METRIC_LINE_RE = re.compile(
    r"^[$~]?\d[\d,]*(?:\.\d+)?(?:\s*/\s*[$~]?\d[\d,]*)?(?:\s*(?:%|k|m|hrs?|hours|million|billion))?\s*$",
    re.I,
)
MILLION_LINE_RE = re.compile(r"^\d+(?:\.\d+)?\s*(?:million|billion)\b", re.I)
MONEY_CALLOUT_RE = re.compile(r"^\$[\d,]")


def compact_text(text: str) -> str:
    return " ".join((text or "").split())


def is_metric_line(text: str) -> bool:
    """True for a KPI token such as 312,000, ~10,000, 13 / 19, or 1 MILLION."""
    t = compact_text(text)
    if not t or len(t.split()) > 4:
        return False
    return bool(METRIC_LINE_RE.match(t) or MILLION_LINE_RE.match(t))


def is_metric_label(text: str) -> bool:
    """Short label that can sit under a KPI number (not another metric)."""
    t = compact_text(text)
    if not t or is_metric_line(t):
        return False
    words = t.split()
    if not (1 <= len(words) <= 10):
        return False
    return not t.endswith((".", "!", "?"))


def looks_like_metric_pairs(items: list[str]) -> bool:
    """True when items alternate KPI number / label (312,000 + students reached)."""
    cleaned = [compact_text(item) for item in items if compact_text(item)]
    if len(cleaned) < 4 or len(cleaned) % 2:
        return False
    for i in range(0, len(cleaned), 2):
        if not is_metric_line(cleaned[i]) or not is_metric_label(cleaned[i + 1]):
            return False
    return True


def pair_metric_label_items(items: list[str]) -> list[list[str]] | None:
    """Group KPI numbers with their labels. None when the list is not that pattern."""
    cleaned = [item for item in items if compact_text(item)]
    if looks_like_metric_pairs(cleaned):
        return [[cleaned[i], cleaned[i + 1]] for i in range(0, len(cleaned), 2)]
    if (
        len(cleaned) >= 5
        and not is_metric_line(cleaned[0])
        and looks_like_metric_pairs(cleaned[1:])
    ):
        rest = cleaned[1:]
        return [[rest[i], rest[i + 1]] for i in range(0, len(rest), 2)]
    if (
        len(cleaned) == 4
        and is_metric_line(cleaned[0])
        and is_metric_line(cleaned[1])
        and is_metric_label(cleaned[2])
        and is_metric_label(cleaned[3])
    ):
        return [[cleaned[0], cleaned[2]], [cleaned[1], cleaned[3]]]
    return None


def is_all_caps_heading(text: str, max_words: int = 8) -> bool:
    t = compact_text(text)
    if not t:
        return False
    words = t.split()
    letters = "".join(ch for ch in t if ch.isalpha())
    return bool(letters) and t == t.upper() and 1 <= len(words) <= max_words


def is_title_like(text: str) -> bool:
    """Slide titles (ALL-CAPS or short labels), not money sentences or captions."""
    t = compact_text(text)
    if not t or is_org_chrome(t):
        return False
    words = t.split()
    if len(words) > 12:
        return False
    if MONEY_CALLOUT_RE.match(t) and len(words) > 4:
        return False
    if t.endswith((".", "!", "?")):
        return False
    if ":" in t and len(words) >= 6:
        return False
    if is_all_caps_heading(t):
        return True
    return 1 <= len(words) <= 8


def looks_like_photo_caption(text: str) -> bool:
    """Sentence-case photo labels ('Marsh survey site'), not INTRODUCING FOREVERGLADES."""
    t = compact_text(text)
    if not t or is_all_caps_heading(t):
        return False
    words = t.split()
    if not (1 <= len(words) <= 16):
        return False
    letters = "".join(ch for ch in t if ch.isalpha())
    if letters and t == t.upper():
        return False
    return True


def title_candidate_score(text: str) -> int:
    """Higher is a better slide title. Money callouts and long sentences lose."""
    t = compact_text(text)
    if not t:
        return -100
    words = t.split()
    n = len(words)
    score = 0
    if 1 <= n <= 8:
        score += 20
    elif n <= 12:
        score += 8
    else:
        score -= 12
    if is_all_caps_heading(t):
        score += 30
    if MONEY_CALLOUT_RE.match(t) or (re.match(r"^[$~]?\d", t) and not is_metric_line(t)):
        score -= 40
    if t.endswith((".", "!", "?")) or (":" in t and n >= 6):
        score -= 16
    if is_org_chrome(t):
        score -= 40
    return score


def is_decorative_icon(shape_or_image: dict[str, Any] | None) -> bool:
    """True for tiny PptxGenJS / UI glyphs — not photographs for photo wells.

    256px transparent PNGs (document, gear, lightning) must not pick PEOPLE
    portraits, Sequence card frames, or a full-bleed Kicker well — even when
    the source scaled the glyph onto a KPI tile. A *low-res* blob that already
    occupies a real photo area (sample marsh ~120px filling the slide) is still
    a photograph. Extreme-aspect grass/rule strips are dividers, not heroes.
    """
    if not shape_or_image:
        return False
    image = shape_or_image.get("image") if isinstance(shape_or_image.get("image"), dict) else shape_or_image
    if not isinstance(image, dict):
        return False
    try:
        area = float(shape_or_image.get("area_ratio") or 0)
    except (TypeError, ValueError):
        area = 0.0
    px = image.get("px") or [0, 0]
    try:
        width, height = int(px[0] or 0), int(px[1] or 0)
    except (TypeError, ValueError, IndexError):
        width = height = 0
    px_max = max(width, height) if width and height else 0
    try:
        nbytes = int(image.get("bytes") or 0)
    except (TypeError, ValueError):
        nbytes = 0
    ext = str(image.get("ext") or image.get("content_type") or "").lower()
    if width and height:
        aspect = width / max(1, height)
        if (aspect >= 8.0 or aspect <= 0.125) and 0 < nbytes < 40000 and "png" in ext:
            return True
    # Stretched thumbnail filling a hero well (120px marsh @ 46% area) is a photo.
    if px_max and px_max <= 160 and area >= 0.15:
        return False
    if px_max and px_max <= 256:
        return True
    if px_max <= 512 and 0 < nbytes < 20000 and "png" in ext:
        if area >= 0.15:
            return False
        return True
    return False


def is_screenshot_blob(blob: bytes, ext: str, aspect: float) -> bool:
    if ext.lower().lstrip(".") != "png":
        return False
    if not (1.2 <= aspect <= 2.2):
        return False
    try:
        from PIL import Image
        from collections import Counter
    except Exception:
        return False
    im = Image.open(io.BytesIO(blob)).convert("RGB")
    im.thumbnail((80, 80))
    counts = Counter(im.getdata())
    if not counts:
        return False
    _color, n = counts.most_common(1)[0]
    return n / max(1, sum(counts.values())) >= 0.18


def apply_line(shape, hex_value: str, width_pt: float) -> None:
    shape.line.color.rgb = rgb(hex_value)
    shape.line.width = Pt(width_pt)


def apply_round_rect(shape) -> None:
    sp_pr = shape._element.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = shape._element.find(qn("p:pic"))
        if sp_pr is None:
            return
        sp_pr = sp_pr.find(qn("p:spPr"))
        if sp_pr is None:
            return
    geom = sp_pr.find(qn("a:prstGeom"))
    if geom is None:
        geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
    geom.set("prst", "roundRect")
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    gd = av.find(qn("a:gd"))
    if gd is None:
        gd = etree.SubElement(av, qn("a:gd"))
    gd.set("name", "adj")
    gd.set("fmla", "val 16667")


def strip_effect_list(shape) -> None:
    for node in shape._element.findall(f".//{{{A_NS}}}effectLst"):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)


def is_footer_shape(shape, text: str, slide_width: int, slide_height: int, left: int, top: int, width: int, height: int) -> bool:
    ph = placeholder_type(shape)
    if ph in FOOTER_PH:
        return True
    blob = (text or "").strip()
    if FOOTER_TEXT_RE.search(blob):
        return True
    bottom = top + height
    words = len(blob.split())
    if bottom >= int(slide_height * 0.94) and 0 < words <= 8:
        return True
    return False


def is_edge_logo(area: int, slide_area: int, left: int, top: int, width: int, height: int, slide_w: int, slide_h: int) -> bool:
    if area > slide_area * 0.08:
        return False
    cx = left + width / 2
    cy = top + height / 2
    band_x = slide_w * 0.10
    band_y = slide_h * 0.10
    return cx <= band_x or cx >= slide_w - band_x or cy <= band_y or cy >= slide_h - band_y


def classify_layout_role(name: str, placeholders: list[dict[str, Any]]) -> str:
    lowered = (name or "").lower()
    types = {(ph.get("type_name") or "").upper() for ph in placeholders}
    names = " ".join((ph.get("name") or "") for ph in placeholders).lower()
    if any(token in lowered for token in ("thank", "closing", "contact", "q&a", "q and a", "end")):
        return "CLOSING"
    if "blank" in lowered:
        return "BLANK"
    if any(token in lowered for token in ("picture", "image", "photo")) or "PICTURE" in types:
        return "PICTURE"
    if any(token in lowered for token in ("two", "comparison", "2 col", "2-col")):
        return "TWO_CONTENT"
    if "section" in lowered or "divider" in lowered or "chapter" in lowered:
        return "SECTION"
    if "title only" in lowered:
        return "TITLE_ONLY"
    if any(token in lowered for token in ("cover",)) or (
        "CENTER_TITLE" in types and "SUBTITLE" in types
    ):
        return "TITLE"
    if lowered.strip() in {"title slide", "title"}:
        return "TITLE"
    body_like = [ph for ph in placeholders if (ph.get("type_name") or "").upper() in {"BODY", "OBJECT", "VERTICAL_BODY"}]
    title_like = [ph for ph in placeholders if (ph.get("type_name") or "").upper() in {"TITLE", "CENTER_TITLE", "VERTICAL_TITLE"}]
    if len(body_like) >= 2 and title_like:
        return "TWO_CONTENT"
    if "title and content" in lowered or ("content" in lowered and "two" not in lowered):
        return "TITLE_BODY"
    if title_like and len(body_like) == 1:
        return "TITLE_BODY"
    if title_like and not body_like:
        return "TITLE_ONLY"
    if not placeholders:
        return "BLANK"
    return "TITLE_BODY"


def resolve_layout_map(layouts: list[dict[str, Any]]) -> dict[str, int]:
    by_role: dict[str, int] = {}
    for layout in layouts:
        role = layout.get("role")
        if role in ROLES and role not in by_role:
            by_role[role] = int(layout["index"])
    mapping: dict[str, int] = {}
    for role in ROLES:
        if role in by_role:
            mapping[role] = by_role[role]
            continue
        for fallback in ROLE_FALLBACKS.get(role, ["TITLE_ONLY", "BLANK"]):
            if fallback in by_role:
                mapping[role] = by_role[fallback]
                break
        else:
            mapping[role] = layouts[0]["index"] if layouts else 0
    return mapping


def layout_index_for_role(tokens: dict[str, Any], role: str) -> int:
    layout_map = tokens.get("layout_map") or {}
    if role in layout_map:
        return int(layout_map[role])
    for fallback in [role, *ROLE_FALLBACKS.get(role, []), "TITLE_ONLY", "BLANK"]:
        if fallback in layout_map:
            return int(layout_map[fallback])
    return 0


def paragraphs_from_shape(shape) -> list[dict[str, Any]]:
    if not getattr(shape, "has_text_frame", False):
        return []
    items = []
    for paragraph in shape.text_frame.paragraphs:
        text = paragraph.text or ""
        if not text.strip() and not paragraph.runs:
            continue
        pPr = paragraph._p.find(qn("a:pPr"))
        numbered = False
        bullet = False
        if pPr is not None:
            numbered = pPr.find(qn("a:buAutoNum")) is not None
            bullet = pPr.find(qn("a:buChar")) is not None or numbered
            if pPr.find(qn("a:buNone")) is not None:
                bullet = False
        align = str(paragraph.alignment) if paragraph.alignment is not None else "left"
        runs = []
        for run in paragraph.runs:
            address = None
            try:
                address = run.hyperlink.address
            except Exception:
                address = None
            runs.append(
                {
                    "text": run.text or "",
                    "bold": bool(run.font.bold) if run.font.bold is not None else None,
                    "italic": bool(run.font.italic) if run.font.italic is not None else None,
                    "underline": bool(run.font.underline) if run.font.underline else False,
                    "hyperlink": address,
                }
            )
        if not runs:
            runs = [{"text": text, "bold": None, "italic": None, "underline": False, "hyperlink": None}]
        items.append(
            {
                "text": text,
                "level": int(paragraph.level or 0),
                "alignment": normalize_alignment(align),
                "bullet": bullet or bool(MANUAL_BULLET_RE.match(text)),
                "numbered": numbered or bool(MANUAL_NUMBER_RE.match(text)),
                "runs": runs,
            }
        )
    return items


def concatenated_text(paragraphs: list[dict[str, Any]]) -> str:
    return "\n".join(p.get("text") or "" for p in paragraphs)


def normalize_for_parity(text: str) -> str:
    cleaned = text.replace("\u00a0", " ")
    cleaned = re.sub(r"[•●▪◦○■□–—\-]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = re.sub(r"\b(page|slide)\s*\d+\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b\d{1,3}\b$", " ", cleaned)
    return cleaned.strip().lower()


SUPPORTED_CHARTS = {
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    XL_CHART_TYPE.COLUMN_STACKED,
    XL_CHART_TYPE.COLUMN_STACKED_100,
    XL_CHART_TYPE.BAR_CLUSTERED,
    XL_CHART_TYPE.BAR_STACKED,
    XL_CHART_TYPE.BAR_STACKED_100,
    XL_CHART_TYPE.LINE,
    XL_CHART_TYPE.LINE_MARKERS,
    XL_CHART_TYPE.PIE,
    XL_CHART_TYPE.PIE_EXPLODED,
    XL_CHART_TYPE.AREA,
    XL_CHART_TYPE.AREA_STACKED,
    XL_CHART_TYPE.DOUGHNUT,
    XL_CHART_TYPE.RADAR,
    XL_CHART_TYPE.RADAR_MARKERS,
    XL_CHART_TYPE.RADAR_FILLED,
}

CHART_FLATTEN = {
    XL_CHART_TYPE.THREE_D_COLUMN: XL_CHART_TYPE.COLUMN_CLUSTERED,
    XL_CHART_TYPE.THREE_D_COLUMN_CLUSTERED: XL_CHART_TYPE.COLUMN_CLUSTERED,
    XL_CHART_TYPE.THREE_D_COLUMN_STACKED: XL_CHART_TYPE.COLUMN_STACKED,
    XL_CHART_TYPE.THREE_D_BAR_CLUSTERED: XL_CHART_TYPE.BAR_CLUSTERED,
    XL_CHART_TYPE.THREE_D_LINE: XL_CHART_TYPE.LINE,
    XL_CHART_TYPE.THREE_D_PIE: XL_CHART_TYPE.PIE,
    XL_CHART_TYPE.THREE_D_AREA: XL_CHART_TYPE.AREA,
}


def enum_name(value: Any) -> str:
    """Normalize python-pptx enum labels ('CENTER_TITLE (3)' → 'CENTER_TITLE')."""
    if value is None:
        return ""
    name = getattr(value, "name", None)
    if isinstance(name, str) and name and name.isidentifier():
        return name
    raw = str(value).split(".")[-1].strip()
    return raw.split("(")[0].strip()


def extract_chart(shape) -> dict[str, Any]:
    info: dict[str, Any] = {
        "type": None,
        "supported": False,
        "external": False,
        "categories": [],
        "series": [],
        "has_data_labels": False,
        "legend_position": None,
        "axis_titles": {},
        "number_format": None,
        "error": None,
    }
    try:
        if chart_is_external(shape):
            info["external"] = True
            info["error"] = "external_xlsx"
            return info
        chart = shape.chart
        raw_type = chart.chart_type
        mapped = CHART_FLATTEN.get(raw_type, raw_type)
        info["type"] = enum_name(mapped)
        info["raw_type"] = enum_name(raw_type)
        info["supported"] = mapped in SUPPORTED_CHARTS
        plot = chart.plots[0] if chart.plots else None
        if plot is not None:
            try:
                info["has_data_labels"] = bool(plot.has_data_labels)
            except Exception:
                pass
            try:
                cats = []
                for cat in plot.categories:
                    cats.append(str(getattr(cat, "label", cat)))
                info["categories"] = cats
            except Exception:
                info["categories"] = []
        try:
            if chart.has_legend:
                info["legend_position"] = enum_name(chart.legend.position)
        except Exception:
            pass
        series_out = []
        try:
            for series in chart.series:
                values = []
                try:
                    values = [float(v) if v is not None else 0.0 for v in (series.values or [])]
                except Exception:
                    values = []
                series_out.append({"name": series.name or "Series", "values": values})
        except Exception as exc:
            info["error"] = str(exc)
        info["series"] = series_out
        if not series_out or not info["categories"]:
            # Some charts expose values only; still try to rebuild if series exist.
            if not series_out:
                info["supported"] = False
        try:
            if chart.category_axis and chart.category_axis.has_title:
                info["axis_titles"]["category"] = chart.category_axis.axis_title.text_frame.text
        except Exception:
            pass
        try:
            if chart.value_axis and chart.value_axis.has_title:
                info["axis_titles"]["value"] = chart.value_axis.axis_title.text_frame.text
        except Exception:
            pass
    except Exception as exc:
        info["error"] = str(exc)
        info["supported"] = False
    return info


def add_rebuilt_chart(slide, chart_info: dict[str, Any], box: dict[str, int], tokens: dict[str, Any], brand: dict[str, Any]):
    from pptx.chart.data import CategoryChartData

    type_name = chart_info.get("type") or "COLUMN_CLUSTERED"
    try:
        chart_type = getattr(XL_CHART_TYPE, type_name)
    except AttributeError:
        chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
    if chart_type not in SUPPORTED_CHARTS:
        chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
    data = CategoryChartData()
    categories = chart_info.get("categories") or []
    series_list = chart_info.get("series") or []
    if not categories and series_list:
        n = max(len(s.get("values") or []) for s in series_list)
        categories = [str(i + 1) for i in range(n)]
    data.categories = categories
    for series in series_list:
        values = list(series.get("values") or [])
        if len(values) < len(categories):
            values = values + [0.0] * (len(categories) - len(values))
        data.add_series(series.get("name") or "Series", tuple(values[: len(categories)]))
    frame = slide.shapes.add_chart(
        chart_type,
        int(box["left"]),
        int(box["top"]),
        int(box["width"]),
        int(box["height"]),
        data,
    )
    chart = frame.chart
    palette = brand.get("chart_palette") or [
        theme_color(tokens, f"accent{i}") for i in range(1, 7)
    ]
    try:
        for i, series in enumerate(chart.series):
            color = hex_color(palette[i % len(palette)])
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = rgb(color)
            try:
                series.format.line.fill.background()
            except Exception:
                pass
    except Exception:
        pass
    try:
        chart.has_legend = bool(chart_info.get("legend_position")) or len(series_list) > 1
        if chart.has_legend:
            pos = (chart_info.get("legend_position") or "BOTTOM").upper()
            mapping = {
                "BOTTOM": XL_LEGEND_POSITION.BOTTOM,
                "TOP": XL_LEGEND_POSITION.TOP,
                "LEFT": XL_LEGEND_POSITION.LEFT,
                "RIGHT": XL_LEGEND_POSITION.RIGHT,
                "CORNER": XL_LEGEND_POSITION.CORNER,
            }
            chart.legend.position = mapping.get(pos, XL_LEGEND_POSITION.BOTTOM)
            chart.legend.include_in_layout = False
    except Exception:
        pass
    try:
        plot = chart.plots[0]
        plot.has_data_labels = bool(chart_info.get("has_data_labels"))
    except Exception:
        pass
    minor = theme_font(tokens, "minor")
    dk1 = theme_color(tokens, "dk1")
    try:
        chart.font.size = Pt(11)
        chart.font.name = minor
        chart.font.color.rgb = rgb(dk1)
    except Exception:
        pass
    try:
        axis = chart.value_axis
        axis.has_major_gridlines = True
        axis.has_minor_gridlines = False
        axis.major_gridlines.format.line.color.rgb = rgb(tint(theme_color(tokens, "dk2"), 0.75))
        axis.major_gridlines.format.line.width = Pt(0.5)
    except Exception:
        pass
    titles = chart_info.get("axis_titles") or {}
    try:
        if titles.get("category"):
            chart.category_axis.has_title = True
            chart.category_axis.axis_title.text_frame.paragraphs[0].text = titles["category"]
        if titles.get("value"):
            chart.value_axis.has_title = True
            chart.value_axis.axis_title.text_frame.paragraphs[0].text = titles["value"]
    except Exception:
        pass
    return frame


def extract_smartart_text(prs_path: Path) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    try:
        with zipfile.ZipFile(prs_path) as zf:
            names = [n for n in zf.namelist() if n.startswith("ppt/diagrams/data") and n.endswith(".xml")]
            for name in names:
                root = etree.fromstring(zf.read(name))
                for pt in root.findall(f".//{{{DGM_NS}}}pt"):
                    texts = [t.text or "" for t in pt.findall(f".//{{{A_NS}}}t") if t.text]
                    if not texts:
                        continue
                    level = 0
                    pr_set = pt.find(f"{{{DGM_NS}}}prSet")
                    if pr_set is not None:
                        for attr in ("lvl", "level", "custT"):
                            if pr_set.get(attr):
                                try:
                                    level = int(pr_set.get(attr))
                                except ValueError:
                                    pass
                    items.append({"text": " ".join(texts).strip(), "level": level, "runs": [{"text": " ".join(texts).strip()}]})
    except Exception:
        return items
    return items


def extract_table(shape) -> dict[str, Any]:
    table = shape.table
    rows = len(table.rows)
    cols = len(table.columns)
    cells = []
    merges = []
    seen_span: set[tuple[int, int]] = set()
    for r in range(rows):
        row_cells = []
        for c in range(cols):
            cell = table.cell(r, c)
            tc = cell._tc
            grid_span = int(tc.get("gridSpan", "1") or 1)
            row_span = int(tc.get("rowSpan", "1") or 1)
            h_merge = tc.get("hMerge") in {"1", "true"}
            v_merge = tc.get("vMerge") in {"1", "true"}
            align = "left"
            try:
                par_align = cell.text_frame.paragraphs[0].alignment
                align = normalize_alignment(par_align)
            except Exception:
                pass
            paragraphs = paragraphs_from_shape(cell)
            row_cells.append(
                {
                    "text": cell.text or "",
                    "paragraphs": paragraphs,
                    "alignment": align,
                    "grid_span": grid_span,
                    "row_span": row_span,
                    "h_merge": h_merge,
                    "v_merge": v_merge,
                    "bold": any(run.get("bold") for p in paragraphs for run in p.get("runs") or []),
                }
            )
            if grid_span > 1 or row_span > 1:
                if (r, c) not in seen_span:
                    merges.append({"row": r, "col": c, "row_span": row_span, "grid_span": grid_span})
                    seen_span.add((r, c))
        cells.append(row_cells)
    widths = [int(col.width) for col in table.columns]
    heights = [int(row.height) for row in table.rows]
    return {
        "rows": rows,
        "cols": cols,
        "cells": cells,
        "merges": merges,
        "col_widths": widths,
        "row_heights": heights,
    }


def style_table(table, tokens: dict[str, Any], brand: dict[str, Any], *, totals_row: bool = False) -> None:
    clear_table_style(table)
    accent1 = theme_color(tokens, "accent1")
    lt1 = theme_color(tokens, "lt1", "FFFFFF")
    lt2 = theme_color(tokens, "lt2", lt1)
    dk1 = theme_color(tokens, "dk1")
    dk2 = theme_color(tokens, "dk2")
    if hex_color(lt2) == hex_color(lt1):
        band = tint(accent1, 0.90)
    else:
        band = lt2
    border = tint(dk2, 0.70)
    rows = len(table.rows)
    cols = len(table.columns)
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            set_cell_margins(cell)
            try:
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
            clear_cell_border(cell, "L")
            clear_cell_border(cell, "R")
            if r == 0:
                fill_cell(cell, accent1)
                set_cell_border(cell, "B", 0.5, border)
                set_cell_border(cell, "T", 0.5, border)
            else:
                fill_cell(cell, lt1 if r % 2 == 1 else band)
                set_cell_border(cell, "T", 0.5, border)
                set_cell_border(cell, "B", 0.5, border)
            if totals_row and r == rows - 1:
                set_cell_border(cell, "T", 1.0, dk1)


def write_table_cell_text(
    cell,
    cell_info: dict[str, Any],
    tokens: dict[str, Any],
    brand: dict[str, Any],
    *,
    header: bool,
    totals: bool,
    font_pt: int,
    curly_quotes: bool,
) -> None:
    paragraphs = cell_info.get("paragraphs") or [
        {"text": cell_info.get("text") or "", "level": 0, "runs": [{"text": cell_info.get("text") or ""}]}
    ]
    color = theme_color(tokens, "lt1") if header else theme_color(tokens, "dk1")
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = True
    align = cell_info.get("alignment") or "left"
    if header:
        align = "center"
    elif looks_number(cell_info.get("text") or ""):
        align = "right"
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = pp_align(align)
        runs = para.get("runs") or [{"text": para.get("text") or ""}]
        for run_info in runs:
            run = p.add_run()
            run.text = apply_quote_style(run_info.get("text") or "", curly_quotes)
            apply_run_style(
                run,
                tokens,
                brand,
                font_kind="minor",
                size_pt=12 if header else font_pt,
                color_hex=color,
                bold=True if header or totals else bool(run_info.get("bold")),
                italic=bool(run_info.get("italic")),
                hyperlink=run_info.get("hyperlink"),
            )


def is_totals_row(row: list[dict[str, Any]]) -> bool:
    if not row:
        return False
    first = (row[0].get("text") or "").strip().lower()
    if re.search(r"\b(total|sum)\b", first):
        return True
    return all(cell.get("bold") for cell in row if (cell.get("text") or "").strip())


def has_timing(slide) -> bool:
    try:
        return slide._element.find(f"{{{P_NS}}}timing") is not None
    except Exception:
        return False


def slide_hidden(slide) -> bool:
    try:
        return slide._element.get("show") == "0"
    except Exception:
        return False


def set_slide_hidden(slide, hidden: bool) -> None:
    if hidden:
        slide._element.set("show", "0")
    elif "show" in slide._element.attrib:
        del slide._element.attrib["show"]


def extract_sections(prs) -> list[dict[str, Any]]:
    sections = []
    try:
        root = prs.element
        for section in root.findall(f".//{{{P14_NS}}}section"):
            name = section.get("name") or ""
            ids = [int(sld.get("id")) for sld in section.findall(f"{{{P14_NS}}}sldId") if sld.get("id")]
            sections.append({"name": name, "sld_ids": ids})
    except Exception:
        return sections
    return sections


def recreate_sections(prs, sections: list[dict[str, Any]], slide_id_order: list[int]) -> None:
    """Best-effort recreation of p14 section markers grouped by original names."""
    if not sections:
        return
    nsmap = {None: P_NS, "p14": P14_NS}
    presentation = prs.element
    ext_lst = presentation.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(presentation, qn("p:extLst"))
    uri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
    for ext in list(ext_lst):
        if ext.get("uri") == uri:
            ext_lst.remove(ext)
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", uri)
    lst = etree.SubElement(ext, f"{{{P14_NS}}}sectionLst")
    sld_id_lst = prs.slides._sldIdLst
    current_ids = [int(sld.get("id")) for sld in sld_id_lst]
    # Map old grouping counts onto new ids in order.
    cursor = 0
    for i, section in enumerate(sections):
        count = len(section.get("sld_ids") or [])
        chunk = current_ids[cursor : cursor + count] if count else []
        if i == len(sections) - 1:
            chunk = current_ids[cursor:]
        cursor += count
        node = etree.SubElement(lst, f"{{{P14_NS}}}section")
        node.set("name", section.get("name") or f"Section {i + 1}")
        node.set("id", f"{0xFFFFFFFF - i:08X}")
        inner = etree.SubElement(node, f"{{{P14_NS}}}sldIdLst")
        for sid in chunk:
            etree.SubElement(inner, f"{{{P14_NS}}}sldId").set("id", str(sid))


def delete_shape(shape) -> None:
    el = shape._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def unused_placeholder(shape) -> bool:
    if not getattr(shape, "is_placeholder", False):
        return False
    if not getattr(shape, "has_text_frame", False):
        return False
    text = (shape.text_frame.text or "").strip()
    return not text


def theme_from_part_xml(xml_bytes: bytes) -> dict[str, Any]:
    root = etree.fromstring(xml_bytes)
    colors: dict[str, str] = {}
    scheme = root.find(f".//{{{A_NS}}}clrScheme")
    names = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
    if scheme is not None:
        for name in names:
            node = scheme.find(f"{{{A_NS}}}{name}")
            if node is None:
                continue
            srgb = node.find(f".//{{{A_NS}}}srgbClr")
            sys_clr = node.find(f".//{{{A_NS}}}sysClr")
            if srgb is not None and srgb.get("val"):
                colors[name] = hex_color(srgb.get("val"))
            elif sys_clr is not None:
                colors[name] = hex_color(sys_clr.get("lastClr") or ("000000" if name.startswith("dk") else "FFFFFF"))
    fonts = {"major": "Calibri", "minor": "Calibri", "major_ea": "", "minor_ea": "", "major_cs": "", "minor_cs": ""}
    major = root.find(f".//{{{A_NS}}}majorFont")
    minor = root.find(f".//{{{A_NS}}}minorFont")
    if major is not None:
        latin = major.find(f"{{{A_NS}}}latin")
        ea = major.find(f"{{{A_NS}}}ea")
        cs = major.find(f"{{{A_NS}}}cs")
        if latin is not None:
            fonts["major"] = latin.get("typeface") or fonts["major"]
        if ea is not None:
            fonts["major_ea"] = ea.get("typeface") or ""
        if cs is not None:
            fonts["major_cs"] = cs.get("typeface") or ""
    if minor is not None:
        latin = minor.find(f"{{{A_NS}}}latin")
        ea = minor.find(f"{{{A_NS}}}ea")
        cs = minor.find(f"{{{A_NS}}}cs")
        if latin is not None:
            fonts["minor"] = latin.get("typeface") or fonts["minor"]
        if ea is not None:
            fonts["minor_ea"] = ea.get("typeface") or ""
        if cs is not None:
            fonts["minor_cs"] = cs.get("typeface") or ""
    return {"colors": colors, "fonts": fonts}


def allowed_fonts(tokens: dict[str, Any], brand: dict[str, Any]) -> set[str]:
    fonts = tokens.get("theme", {}).get("fonts") or {}
    allowed = {fonts.get("major"), fonts.get("minor"), fonts.get("major_ea"), fonts.get("minor_ea"), fonts.get("major_cs"), fonts.get("minor_cs")}
    allowed.update(brand.get("fonts") or [])
    if brand.get("east_asian_font"):
        allowed.add(brand["east_asian_font"])
    if brand.get("complex_font"):
        allowed.add(brand["complex_font"])
    return {f for f in allowed if f}


def allowed_colors(tokens: dict[str, Any], brand: dict[str, Any]) -> set[str]:
    colors = set((tokens.get("theme", {}).get("colors") or {}).values())
    colors.update(hex_color(c) for c in (brand.get("color_whitelist") or []))
    colors.update(hex_color(c) for c in (brand.get("chart_palette") or []))
    base = list(colors)
    for c in base:
        for pct in (0.12, 0.10, 0.25, 0.30, 0.50, 0.70, 0.75, 0.90):
            colors.add(tint(c, pct))
    colors.add("FFFFFF")
    colors.add("000000")
    return {hex_color(c) for c in colors if c}


def uses_curly_quotes(tokens: dict[str, Any]) -> bool:
    sample = tokens.get("sample_title_slide_fonts") or {}
    blob = json.dumps(sample) + json.dumps(tokens.get("quote_probe") or "")
    return ("“" in blob) or ("”" in blob) or bool(tokens.get("curly_quotes"))
