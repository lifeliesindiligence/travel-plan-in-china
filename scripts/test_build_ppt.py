"""Offline regression checks: python -m unittest discover -s scripts -p 'test_*.py'."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation
from pptx.oxml.ns import qn

import build_ppt as renderer

ROOT = Path(__file__).resolve().parents[1]


def contrast(a, b):
    def luminance(value):
        channels = [int(value[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        linear = [c / 12.92 if c <= .04045 else ((c + .055) / 1.055) ** 2.4 for c in channels]
        return sum(c * w for c, w in zip(linear, (.2126, .7152, .0722)))
    high, low = sorted((luminance(a), luminance(b)), reverse=True)
    return (high + .05) / (low + .05)


class RendererTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.plan = json.loads((ROOT / "examples/example-plan.json").read_text(encoding="utf-8"))

    def build(self, plan=None, **kwargs):
        output = self.base / "test.pptx"
        renderer.build(plan or self.plan, self.base, output, **kwargs)
        with zipfile.ZipFile(output) as archive:
            self.assertIsNone(archive.testzip())
        return Presentation(output)

    def test_text_palette_has_normal_text_contrast(self):
        for foreground, background in (
            (renderer.WHITE, renderer.NAVY), (renderer.WHITE, renderer.TEAL),
            (renderer.WHITE, renderer.ORANGE_DARK), (renderer.WHITE, renderer.CLOSED),
            (renderer.ORANGE_DARK, renderer.WHITE), (renderer.MUTED, renderer.PAPER),
            (renderer.GOLD, renderer.NAVY),
        ):
            with self.subTest(foreground=foreground, background=background):
                self.assertGreaterEqual(contrast(foreground, background), 4.5)

    def test_fonts_are_declared_for_every_run(self):
        with patch.dict(os.environ, {}, clear=True):
            prs = self.build()
        runs = [run for slide in prs.slides for shape in slide.shapes if shape.has_text_frame
                for paragraph in shape.text_frame.paragraphs for run in paragraph.runs]
        self.assertTrue(runs)
        for run in runs:
            self.assertEqual(run.font.name, "Microsoft YaHei")
            self.assertEqual(run._r.rPr.find(qn("a:ea")).get("typeface"), "微软雅黑")
            self.assertEqual(run._r.rPr.find(qn("a:cs")).get("typeface"), "Microsoft YaHei")

    def test_explicit_font_overrides_environment_and_resets(self):
        with patch.dict(os.environ, {"TRAVEL_PLAN_FONT": "Source Han Sans SC"}):
            self.build()
            self.assertEqual(renderer.FONT, "Source Han Sans SC")
            prs = self.build(font="PingFang SC")
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            self.assertEqual(run._r.rPr.find(qn("a:ea")).get("typeface"), "PingFang SC")
        with patch.dict(os.environ, {}, clear=True):
            self.build()
            self.assertEqual(renderer.FONT, "Microsoft YaHei")

    def test_pagination_preserves_order_and_capacity(self):
        for capacity in (6, 7):
            for count in range(1, 65):
                items = list(range(count))
                pages = list(renderer.balanced_pages(items, capacity))
                sizes = [len(chunk) for _, chunk in pages]
                self.assertEqual([item for _, chunk in pages for item in chunk], items)
                self.assertLessEqual(max(sizes), capacity)
                self.assertLessEqual(max(sizes) - min(sizes), 1)
        self.assertEqual([len(p) for _, p in renderer.balanced_pages(list(range(10)), 6)], [5, 5])
        self.assertEqual([len(p) for _, p in renderer.balanced_pages(list(range(8)), 7)], [4, 4])

    def test_generated_ten_events_and_eight_sources(self):
        self.plan["itinerary"] = [dict(start=f"{8+i:02d}:00", end=f"{8+i:02d}:30", activity=f"活动{i:02d}") for i in range(10)]
        self.plan["sources"] = [dict(id=f"s{i}", title=f"来源{i:02d}", url=f"https://example.org/{i}", observed_at="2026-11-27") for i in range(8)]
        prs = self.build()
        event_pages, source_pages, links = [], [], []
        for slide in prs.slides:
            labels = [s.text for s in slide.shapes if s.has_text_frame]
            events = [t for t in labels if t.startswith("活动")]
            sources = [t for t in labels if t in {f"来源{i:02d}" for i in range(8)}]
            if events:
                event_pages.append(events)
                time_boxes = [s for s in slide.shapes if s.has_text_frame and "–" in s.text]
                self.assertEqual(len(time_boxes), 5)
                # Short tables must be vertically centered in the six-row area.
                self.assertAlmostEqual(time_boxes[0].top / 914400, 2.03 + .375 + .17, places=3)
            if sources:
                source_pages.append(sources)
            for shape in slide.shapes:
                if shape.has_text_frame:
                    links.extend(run.hyperlink.address for p in shape.text_frame.paragraphs for run in p.runs if run.hyperlink.address)
        self.assertEqual([len(p) for p in event_pages], [5, 5])
        self.assertEqual([len(p) for p in source_pages], [4, 4])
        self.assertEqual(len(set(links)), 8)

    def test_verified_closed_has_red_badge_and_attraction_marker(self):
        self.plan["operation"]["status"] = "verified_closed"
        self.plan["itinerary"] = []
        self.plan["attractions"][0]["operation"] = "verified_closed"
        prs = self.build()
        badge = next(s for s in prs.slides[0].shapes if abs(s.top / 914400 - 5.12) < .01)
        self.assertEqual(str(badge.fill.fore_color.rgb), renderer.CLOSED)
        cards = prs.slides[2]
        self.assertTrue(any(s.fill.type is not None and s.fill.fore_color.type is not None
                            and str(s.fill.fore_color.rgb) == renderer.CLOSED
                            for s in cards.shapes if not s.has_text_frame or not s.text))

    def test_footer_minimum_size(self):
        prs = self.build()
        for slide in list(prs.slides)[1:]:
            for shape in slide.shapes:
                if shape.has_text_frame and shape.top / 914400 >= 7.16:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            self.assertGreaterEqual(run.font.size.pt, 9.5)

    def test_cli_accepts_font_option(self):
        output = self.base / "cli.pptx"
        result = subprocess.run(
            [sys.executable, str(ROOT / "scripts/build_ppt.py"), str(ROOT / "examples/example-plan.json"), str(output), "--font", "Source Han Sans SC"],
            stdin=subprocess.DEVNULL, capture_output=True, text=True, encoding="utf-8",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["font"], "Source Han Sans SC")
        self.assertTrue(output.is_file())


if __name__ == "__main__":
    unittest.main()
