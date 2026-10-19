#!/usr/bin/env python3
"""Regression: leftover prototype dummy must not stack on source copy."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
sys.path.insert(0, str(ROOT))

from pptx import Presentation  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa: E402

from colorways import is_layout_chrome  # noqa: E402
from helpers import blob_hash  # noqa: E402
from test_margins_images import KNIGHT, SAMPLE, TEMPLATE, _convert  # noqa: E402

EMU_PER_INCH = 914400

LEFTOVER_DUMMY = (
    "Section title goes here",
    "A sentence framing what this section sets out to do",
    "Title template",
    "Bullet one",
    "Bullet two",
    "Description text here",
    "PRESENTER NAME",
    "A subtitle line goes here",
    "Restoring America’s Everglades",
    "Restoring America's Everglades",
    "Two or three sentences of body copy",
    "A second short paragraph",
    "Concept one",
    "Concept two",
    "Concept three",
    "One sentence that carries the whole argument",
    "A short line of support underneath",
    "LABEL HERE",
    "One line explaining what the figure counts",
    "COLUMN ONE",
    "Gallery template",
    "One line naming what this photograph shows",
    "Section title here",
    "One line of description for this section",
)


def _iter_text(slide) -> list[str]:
    found: list[str] = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if text:
            found.append(text)
    return found


def _assert_no_dummy(slide, label: str) -> None:
    blob = "\n".join(_iter_text(slide))
    for dummy in LEFTOVER_DUMMY:
        if dummy in blob:
            raise AssertionError(f"{label} still has prototype copy {dummy!r}")


class ChromeClassifyTests(unittest.TestCase):
    def test_keeps_layout_labels_and_index_badges(self):
        self.assertTrue(is_layout_chrome("TWO POINTS"))
        self.assertTrue(is_layout_chrome("KICKER"))
        self.assertTrue(is_layout_chrome("01"))
        self.assertTrue(is_layout_chrome("PRESENTATION TITLE"))
        self.assertFalse(is_layout_chrome("Section title goes here"))
        self.assertFalse(is_layout_chrome("Bullet one"))
        self.assertFalse(is_layout_chrome("PRESENTER NAME · MONTH 2026"))


@unittest.skipUnless(TEMPLATE.exists() and SAMPLE.exists(), "template or sample fixture missing")
class SampleOverlapTests(unittest.TestCase):
    def test_sample_has_no_stacked_prototype_copy(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = _convert(TEMPLATE, SAMPLE, Path(tmp))
            prs = Presentation(str(result["output"]))
            for i, slide in enumerate(prs.slides):
                _assert_no_dummy(slide, f"sample slide {i}")


@unittest.skipUnless(TEMPLATE.exists() and KNIGHT.exists(), "Knight fixture not in this environment")
class KnightOverlapTests(unittest.TestCase):
    def test_knight_dummy_cleared_and_source_in_wells(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = _convert(TEMPLATE, KNIGHT, Path(tmp))
            prs = Presentation(str(result["output"]))
            self.assertEqual(len(prs.slides), 6)

            cover = "\n".join(_iter_text(prs.slides[0]))
            _assert_no_dummy(prs.slides[0], "cover")
            self.assertIn("Knight Foundation Grant", cover)
            self.assertIn("PRESENTATION TITLE", cover)

            section = prs.slides[3]
            texts = _iter_text(section)
            blob = "\n".join(texts)
            _assert_no_dummy(section, "section divider")
            self.assertTrue(any(is_layout_chrome(t) for t in texts), "section number chrome must remain")
            self.assertIn("Expanded Email Summary and Knight Foundation Profile", blob)
            self.assertNotEqual(texts[0], "Expanded Email Summary and Knight Foundation Profile")

            two = prs.slides[2]
            two_blob = "\n".join(_iter_text(two))
            _assert_no_dummy(two, "two points")
            self.assertIn("TWO POINTS", two_blob)
            self.assertIn("Grant Status, Financial Overview, and Key Initiatives", two_blob)
            self.assertIn("Grant Progress and Compliance", two_blob)
            self.assertTrue(any(t.strip() == "01" for t in _iter_text(two)))
            self.assertTrue(any(t.strip() == "02" for t in _iter_text(two)))
            for shape in two.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                text = (shape.text_frame.text or "").strip()
                if text in {"Grant Progress and Compliance", "Key Funded Initiatives"}:
                    self.assertGreater(
                        int(shape.width or 0),
                        int(2.0 * EMU_PER_INCH),
                        "column headings must sit in heading wells, not the 01/02 badges",
                    )

            kicker = "\n".join(_iter_text(prs.slides[4]))
            _assert_no_dummy(prs.slides[4], "kicker")
            self.assertIn("KICKER", kicker)
            self.assertIn("Summary of Jodi Farrell", kicker)
            self.assertIn("Central Coordination Role", kicker)

            profile = "\n".join(_iter_text(prs.slides[5]))
            _assert_no_dummy(prs.slides[5], "profile")
            self.assertIn("TWO POINTS", profile)
            self.assertIn("Mission and Legacy", profile)
            self.assertTrue(any(t.strip() == "01" for t in _iter_text(prs.slides[5])))

            # Unlocked source photos still present (PR #4).
            source_hashes = {
                shape["image"]["hash"]
                for slide in result["manifest"]["slides"]
                for shape in slide["shapes"]
                if shape.get("kind") == "picture" and not shape.get("drop_logo")
            }
            found = set()
            A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
                        continue
                    digest = blob_hash(shape.image.blob)
                    if digest in source_hashes:
                        found.add(digest)
                        self.assertFalse(shape._element.findall(f".//{{{A_NS}}}picLocks"))
            self.assertEqual(found, source_hashes)


if __name__ == "__main__":
    unittest.main(verbosity=2)
