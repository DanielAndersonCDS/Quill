"""Exercise quill.py's docx helpers without PySide6 by extracting them with ast."""
import ast, pathlib, unittest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

SRC = (pathlib.Path(__file__).parent.parent / "quill.py").read_text(encoding="utf-8")


def grab(*names):
    tree, ns = ast.parse(SRC), {"parse_xml": parse_xml, "nsdecls": nsdecls, "qn": qn}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            exec(compile(ast.Module([node], []), "quill", "exec"), ns)
    return ns


class NumberingTests(unittest.TestCase):
    def test_restart_numbering_and_levels(self):
        ns = grab("_new_list_num")
        d = Document()
        ids = [ns["_new_list_num"](d, "List Number") for _ in range(3)]
        self.assertEqual(len(set(ids)), 3)
        self.assertTrue(all(isinstance(i, int) for i in ids))
        for n, nid in enumerate(ids):
            p = d.add_paragraph(style="List Number")
            numPr = p._p.get_or_add_pPr().get_or_add_numPr()
            numPr.get_or_add_numId().val = nid
            numPr.get_or_add_ilvl().val = 0
        for st in ("List Bullet 2", "List Bullet 3", "List Number 2", "List Number 3"):
            d.add_paragraph(style=st)
        d.save("/tmp/num_test.docx")
        d2 = Document("/tmp/num_test.docx")
        nums = d2.part.numbering_part.element.xpath("./w:num")
        self.assertGreaterEqual(len(nums), 3)
        overrides = d2.part.numbering_part.element.xpath("//w:startOverride/@w:val")
        self.assertEqual(overrides.count("1"), 3)

    def test_box_border_placed_before_vertalign(self):
        d = Document()
        run = d.add_paragraph().add_run("x")
        run.font.superscript = True
        rPr = run._r.get_or_add_rPr()
        bdr = parse_xml(f'<w:bdr {nsdecls("w")} w:val="single" w:sz="4" w:space="1" w:color="auto"/>')
        va = rPr.find(qn("w:vertAlign"))
        va.addprevious(bdr)
        tags = [c.tag.split("}")[1] for c in rPr]
        self.assertLess(tags.index("bdr"), tags.index("vertAlign"))


class BidiTests(unittest.TestCase):
    def test_bidi_is_placed_before_spacing_ind_jc(self):
        from docx.shared import Pt
        ns = grab("_set_bidi")
        d = Document()
        p = d.add_paragraph("x")
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.right_indent = Pt(10)
        p.alignment = 1
        ns["_set_bidi"](p)
        ns["_set_bidi"](p)  # idempotent
        tags = [c.tag.split("}")[1] for c in p._p.pPr]
        self.assertEqual(tags.count("bidi"), 1)
        self.assertLess(tags.index("bidi"), tags.index("spacing"))
        self.assertLess(tags.index("bidi"), tags.index("ind"))
        self.assertLess(tags.index("bidi"), tags.index("jc"))


if __name__ == "__main__":
    unittest.main()
