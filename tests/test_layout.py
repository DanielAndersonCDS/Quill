import unittest, tempfile, os
from docx import Document
from docx.oxml.ns import qn
from quill_layout import PageSetup, apply_page_setup, read_page_setup


def roundtrip(ps):
    d = Document()
    d.add_paragraph("body")
    apply_page_setup(d, ps)
    path = os.path.join(tempfile.mkdtemp(), "t.docx")
    d.save(path)
    d2 = Document(path)  # must re-open
    return d2, read_page_setup(d2)


class LayoutRoundTrip(unittest.TestCase):
    def test_defaults(self):
        _, r = roundtrip(PageSetup())
        self.assertEqual((r.size_name, r.orientation, r.margin_name, r.columns), ("A4", "portrait", "Normal", 1))
        self.assertFalse(r.border or r.line_numbers or r.hyphenate)
        self.assertEqual((r.header, r.footer, r.watermark), ("", "", ""))

    def test_everything(self):
        ps = PageSetup(width_mm=215.9, height_mm=279.4, orientation="landscape", top=12.7, bottom=12.7,
                       left=50.8, right=50.8, columns=3, col_gap=10, header="My Header",
                       footer="Page {page} of {pages}", watermark="CONFIDENTIAL", border=True,
                       border_style="double", border_width=2.25, border_color="FF0000",
                       line_numbers=True, hyphenate=True)
        d, r = roundtrip(ps)
        self.assertEqual((r.size_name, r.orientation), ("Letter", "landscape"))
        self.assertEqual(r.margin_name, "Custom")
        self.assertEqual(r.columns, 3)
        self.assertAlmostEqual(r.col_gap, 10, delta=0.1)
        self.assertEqual(r.header, "My Header")
        self.assertEqual(r.footer, "Page {page} of {pages}")
        self.assertEqual(r.watermark, "CONFIDENTIAL")
        self.assertTrue(r.border and r.line_numbers and r.hyphenate)
        self.assertEqual((r.border_style, r.border_color), ("double", "FF0000"))
        self.assertAlmostEqual(r.border_width, 2.25)

    def test_sectpr_child_order(self):
        d, _ = roundtrip(PageSetup(border=True, line_numbers=True, columns=2))
        tags = [c.tag.split("}")[1] for c in d.sections[0]._sectPr]
        order = ["pgSz", "pgMar", "pgBorders", "lnNumType", "cols"]
        idx = [tags.index(t) for t in order]
        self.assertEqual(idx, sorted(idx))

    def test_special_characters_in_watermark(self):
        _, r = roundtrip(PageSetup(watermark='A & "B" <C>'))
        self.assertEqual(r.watermark, 'A & "B" <C>')

    def test_single_column_has_no_num(self):
        d, _ = roundtrip(PageSetup(columns=1))
        self.assertIsNone(d.sections[0]._sectPr.find(qn("w:cols")).get(qn("w:num")))


if __name__ == "__main__":
    unittest.main()
