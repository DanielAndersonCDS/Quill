import pathlib, unittest
from quill_paste.mathconv import convert, plain
from quill_paste.parser import looks_structured, parse

HERE = pathlib.Path(__file__).parent


def P(src, display=True):
    return plain(convert(src, display))


def blocks(name):
    return parse((HERE / name).read_text(encoding="utf-8"))


class MathTests(unittest.TestCase):
    def test_logic_and_sets(self):
        self.assertEqual(P(r"\forall x,y\in A,\quad xRy\Rightarrow yRx"), "∀x, y ∈ A, \u2003xRy ⇒ yRx")
        self.assertEqual(P(r"R^{-1}=\{(y,x)\in B\times A\mid(x,y)\in R\}"),
                         "R−1 = {(y, x) ∈ B × A ∣ (x, y) ∈ R}")
        self.assertEqual(P(r"E:\mathbb Z\rightarrow\mathbb Z"), "E: ℤ → ℤ")
        self.assertEqual(P(r"3\mid(x-3)"), "3 ∣ (x − 3)")

    def test_superscript_and_subscript_flags(self):
        r = convert(r"R^{-1}")
        self.assertTrue(any(x.sup and x.text == "−1" for x in r))
        r = convert(r"1\leq_R2")
        self.assertTrue(any(x.sub and x.text == "R" for x in r))
        self.assertEqual(plain(r), "1 ≤R 2")  # script stays glued to the operator

    def test_negation_and_unary_minus(self):
        self.assertEqual(P(r"5\not E2"), "5E̸2")
        self.assertEqual(P(r"x\notin A"), "x ∉ A")
        self.assertEqual(P(r"\{-3,-1\}"), "{−3, −1}")
        self.assertEqual(P(r"3-1=2"), "3 − 1 = 2")

    def test_comma_spacing_in_sets(self):
        self.assertEqual(P(r"\{0,3,6\}"), "{0, 3, 6}")
        self.assertEqual(P(r"1,000"), "1,000")

    def test_boxed_and_text(self):
        r = convert(r"\boxed{R\text{ is reflexive}}")
        self.assertEqual(plain(r), "R is reflexive")
        self.assertTrue(all(x.box for x in r))

    def test_aligned_environment_makes_line_breaks(self):
        s = P(r"\begin{aligned}R=\{&(a,a),\\&(a,b)\}\end{aligned}")
        self.assertEqual(s, "R = {(a, a),\u2028(a, b)}")

    def test_misc(self):
        self.assertEqual(P(r"\frac{a+b}{c}"), "(a + b)/c")
        self.assertEqual(P(r"x\<y"), "x < y")
        self.assertEqual(P(r"\operatorname{LUB}(\{i,j\})=d"), "LUB({i, j}) = d")
        self.assertIn("\\foo", P(r"\foo x"))  # unknown commands stay visible


class ParseDoc1Tests(unittest.TestCase):
    def setUp(self):
        self.b = blocks("fixture_doc1.txt")
        self.kinds = [b.kind for b in self.b]

    def test_detected(self):
        self.assertTrue(looks_structured((HERE / "fixture_doc1.txt").read_text(encoding="utf-8")))

    def test_headings_levels(self):
        heads = [(b.level, plain(b.runs)) for b in self.b if b.kind == "heading"]
        self.assertEqual(heads[0], (1, "Properties of Relations — Complete 49-Slide Lesson"))
        self.assertIn((2, "Slide 6 — What exactly is a relation?"), heads)
        self.assertIn((3, "1.1 Start with sets"), heads)

    def test_tables(self):
        tables = [b for b in self.b if b.kind == "table"]
        self.assertEqual(len(tables), 2)
        self.assertEqual([plain(c) for c in tables[1].rows[1]], ["2", "2", "2 = 1(2)", "Yes"])
        self.assertTrue(all(r.bold for c in tables[0].rows[0] for r in c))

    def test_display_math_boxed_and_escapes(self):
        maths = [plain(b.runs) for b in self.b if b.kind == "math"]
        self.assertIn("A = {1, 2, 3}", maths)
        self.assertIn("mEn ⇔ m − n is even", maths)
        self.assertIn("xRy ⇔ x < y", maths)
        self.assertTrue(any(r.box for b in self.b if b.kind == "math" for r in b.runs))

    def test_lists_quote_and_entity(self):
        self.assertEqual(self.kinds.count("number"), 2)
        self.assertEqual(self.kinds.count("bullet"), 2)
        self.assertIn("quote", self.kinds)
        para = next(b for b in self.b if plain(b.runs).startswith("Example 1"))
        self.assertEqual(plain(para.runs), "Example 1: Is 4E0?")
        self.assertNotIn("**", plain(para.runs))


class ParseDoc2Tests(unittest.TestCase):
    def setUp(self):
        self.b = blocks("fixture_doc2.txt")

    def test_plain_text_headings(self):
        heads = [(b.level, plain(b.runs)) for b in self.b if b.kind == "heading"]
        self.assertEqual(heads[0][0], 1)
        self.assertIn((2, "Understanding how to approach mathematical definitions"), heads)
        self.assertIn((3, "Proving something is true versus proving something is false"), heads)
        self.assertIn((3, "Hasse diagram of divisibility"), heads)

    def test_period_ended_numbered_lines_stay_lists(self):
        nums = [plain(b.runs) for b in self.b if b.kind == "number"]
        self.assertEqual(len(nums), 2)

    def test_tab_table(self):
        t = next(b for b in self.b if b.kind == "table")
        self.assertEqual(len(t.rows), 3)
        self.assertEqual(plain(t.rows[0][1]), "To prove true")

    def test_multiline_display_math(self):
        maths = [plain(b.runs) for b in self.b if b.kind == "math"]
        self.assertIn("∀x ∈ A, \u2003xRx", maths)
        self.assertIn("R = {(a, a), (b, b),\u2028(a, b), (b, a)}", maths)
        self.assertIn("Maximal = {6}", maths)

    def test_sentences_are_not_headings(self):
        paras = [plain(b.runs) for b in self.b if b.kind == "para"]
        self.assertIn("Therefore:", paras)
        self.assertIn("A line means the lower element divides the higher element.", paras)


class DetectionTests(unittest.TestCase):
    def test_plain_prose_is_left_alone(self):
        self.assertFalse(looks_structured("Hello there, this is a normal sentence.\nAnother line."))
        self.assertFalse(looks_structured("It costs $5 and $10 in total."))

    def test_markup_triggers(self):
        self.assertTrue(looks_structured("- one\n- two\n- three"))
        self.assertTrue(looks_structured("# Title\nbody"))
        self.assertTrue(looks_structured("Name\tAge\nBob\t3"))

    def test_currency_not_math(self):
        r = parse("Total $5 and $10 later")[0].runs
        self.assertEqual(plain(r), "Total $5 and $10 later")

    def test_robust_on_garbage(self):
        for t in ["", "\n\n", "\\[ unclosed", "| a |\n", "**", "$$", "\\begin{aligned}"]:
            parse(t)  # must not raise


if __name__ == "__main__":
    unittest.main()
