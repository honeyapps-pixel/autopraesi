"""Tests: Zeilenumbrüche und frei bearbeiteter Untertitel auf der Thema-Folie.

Ausführen:  cd autopraesi_python && venv/bin/python -m unittest -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

from config import TEMPLATE_PATH  # noqa: E402
from excel_reader import GodiPlanData  # noqa: E402
import presentation_builder as pb  # noqa: E402

THEME_SLIDE = 3  # Template-Folie 4 ("Thema" / "Gottesdienst am XY.XY.XY")


def _data(**kw):
    d = GodiPlanData(date_str="08.11.2026", kirchenkalender="Drittletzter Sonntag des Kirchenjahres")
    for k, v in kw.items():
        setattr(d, k, v)
    return d


def _shape(slide, name):
    return next(s for s in slide.shapes if s.name == name)


class ThemeLineBreakTest(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation(TEMPLATE_PATH)
        self.slide = self.prs.slides[THEME_SLIDE]

    def test_newline_becomes_line_break(self):
        data = _data(theme="Zwischen Hoffnung und Erfüllung –\nJesus macht den Unterschied")
        pb._fill_slide_simple(self.slide, data)
        title = _shape(self.slide, "Titel 1")
        paras = title.text_frame.paragraphs
        self.assertEqual(len(paras), 1)
        p = paras[0]._p
        self.assertEqual(len(p.findall(qn("a:br"))), 1)
        self.assertEqual([r.text for r in paras[0].runs],
                         ["Zwischen Hoffnung und Erfüllung –", "Jesus macht den Unterschied"])
        # Beide Runs gleich formatiert (Schriftgröße bleibt erhalten)
        sizes = {r.font.size for r in paras[0].runs}
        self.assertEqual(len(sizes), 1)

    def test_single_line_unchanged(self):
        data = _data(theme="Gottes Liebe erkennen")
        pb._fill_slide_simple(self.slide, data)
        title = _shape(self.slide, "Titel 1")
        self.assertEqual(title.text_frame.text, "Gottes Liebe erkennen")
        self.assertEqual(title.text_frame.paragraphs[0]._p.findall(qn("a:br")), [])

    def test_layout_font_size_reaches_all_lines(self):
        data = _data(theme="Zeile eins\nZeile zwei\nZeile drei")
        pb._fill_slide_simple(self.slide, data)
        pb._apply_theme_text_layout(self.prs, {"x": 5, "y": 5, "w": 90, "h": 30, "fontSize": 5})
        runs = _shape(self.slide, "Titel 1").text_frame.paragraphs[0].runs
        self.assertEqual(len(runs), 3)
        self.assertTrue(all(r.font.size.pt == 50 for r in runs))


class SubtitleOverrideTest(unittest.TestCase):
    def test_subtitle_override_on_all_theme_slides(self):
        prs = Presentation(TEMPLATE_PATH)
        data = _data(theme="Thema")
        for slide in prs.slides:
            pb._fill_slide_simple(slide, data)
        pb._apply_subtitle_override(prs, "Gottesdienst am 08.11.2026\nmit Taufe")
        sub = _shape(prs.slides[THEME_SLIDE], "CustomShape 1")
        self.assertEqual([p.text for p in sub.text_frame.paragraphs],
                         ["Gottesdienst am 08.11.2026", "mit Taufe"])
        # Formatierung des ersten Runs übernommen
        sizes = {p.runs[0].font.size for p in sub.text_frame.paragraphs}
        self.assertEqual(len(sizes), 1)

    def test_set_shape_lines_handles_text_without_marker(self):
        prs = Presentation(TEMPLATE_PATH)
        sub = _shape(prs.slides[THEME_SLIDE], "CustomShape 1")
        pb._set_shape_lines(sub, "Familiengottesdienst")
        self.assertEqual(sub.text_frame.text, "Familiengottesdienst")


if __name__ == "__main__":
    unittest.main()
