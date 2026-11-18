import hashlib
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = "ppt/presentation.xml"


class OrderTests(unittest.TestCase):
    def setUp(self):
        self.scratch = TemporaryDirectory()
        self.addCleanup(self.scratch.cleanup)
        self.demo = Path(self.scratch.name) / "fixtures"
        subprocess.run([sys.executable, ROOT / "examples/make_decks.py", self.demo],
                       check=True, capture_output=True)
        self.mosaic = self.demo / "mosaic.pptx"
        self.output = self.demo / "ordered.pptx"

    def run_order(self, source, *order, output=None):
        return subprocess.run([sys.executable, "-m", "pptx_order", source, "--order",
                               *map(str, order), "-o", output or self.output],
                              cwd=ROOT, text=True, capture_output=True)

    def test_reorders_native_chart_and_preserves_package(self):
        initial_digest = hashlib.sha256(self.mosaic.read_bytes()).hexdigest()
        result = self.run_order(self.mosaic, 3, 1, 2)
        self.assertEqual(result.returncode, 0, result.stderr)
        deck = Presentation(self.output)
        self.assertEqual([slide.shapes.title.text for slide in deck.slides],
                         ["Lay the border", "Pick three tiles", "Count the colors"])
        chart = next(shape.chart for shape in deck.slides[2].shapes if shape.has_chart)
        self.assertEqual(tuple(chart.series[0].values), (5.0, 3.0, 2.0))
        self.assertEqual(hashlib.sha256(self.mosaic.read_bytes()).hexdigest(), initial_digest)
        with ZipFile(self.mosaic) as original, ZipFile(self.output) as ordered:
            self.assertCountEqual(original.namelist(), ordered.namelist())
            for name in original.namelist():
                if name != MANIFEST:
                    self.assertEqual(ordered.read(name), original.read(name), name)

    def test_two_slide_deck(self):
        result = self.run_order(self.demo / "lantern.pptx", 2, 1)
        self.assertEqual(result.returncode, 0, result.stderr)
        deck = Presentation(self.output)
        self.assertEqual([slide.shapes.title.text for slide in deck.slides],
                         ["Join the ends", "Fold the shade"])

    def test_requires_complete_permutation(self):
        for order in ((1, 2), (1, 2, 2), (1, 2, 4), (3, 2, 0)):
            with self.subTest(order=order):
                result = self.run_order(self.mosaic, *order)
                self.assertEqual(result.returncode, 2)
                self.assertIn("exactly once", result.stderr)
                self.assertFalse(self.output.exists())

    def test_refuses_output_collisions(self):
        self.output.write_bytes(b"keep this file")
        result = self.run_order(self.mosaic, 1, 2, 3)
        self.assertEqual(result.returncode, 2)
        self.assertIn("already exists", result.stderr)
        self.assertEqual(self.output.read_bytes(), b"keep this file")
        result = self.run_order(self.mosaic, 3, 2, 1, output=self.mosaic)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(Presentation(self.mosaic).slides[0].shapes.title.text, "Pick three tiles")

    def test_damaged_input(self):
        self.mosaic.write_bytes(b"not a presentation")
        result = self.run_order(self.mosaic, 1)
        self.assertEqual(result.returncode, 2)
        self.assertIn("readable PPTX", result.stderr)
        self.assertFalse(self.output.exists())

    def test_refuses_sections(self):
        sectioned = self.demo / "sections.pptx"
        with ZipFile(self.mosaic) as original, ZipFile(sectioned, "w") as changed:
            for name in original.namelist():
                data = original.read(name)
                if name == MANIFEST:
                    root = etree.fromstring(data)
                    etree.SubElement(root, "{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst")
                    data = etree.tostring(root)
                changed.writestr(name, data)
        result = self.run_order(sectioned, 2, 3, 1)
        self.assertEqual(result.returncode, 2)
        self.assertIn("sectioned decks", result.stderr)
        self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
