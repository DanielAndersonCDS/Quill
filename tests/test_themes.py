import json, pathlib, tempfile, unittest
from quill_themes import load_themes, palette_from_tokens, mix

ROOT = pathlib.Path(__file__).parent.parent
KEYS = ("rosewater pink mauve red peach yellow green teal blue lavender text subtext1 subtext0 "
        "overlay0 surface2 surface1 surface0 base mantle crust accent_bg sapphire sky maroon "
        "flamingo accent_text").split()


class ThemeTests(unittest.TestCase):
    def test_shipped_themes_have_every_key(self):
        themes, problems = load_themes(ROOT / "themes")
        self.assertEqual(problems, [])
        self.assertEqual(sorted(themes), ["bloomberg", "nord"])  # mocha/latte are built in
        for pal in themes.values():
            for k in KEYS:
                self.assertIn(k, pal)

    def test_dark_detection(self):
        themes, _ = load_themes(ROOT / "themes")
        self.assertTrue(themes["nord"]["dark"])

    def test_bad_files_reported_not_fatal(self):
        with tempfile.TemporaryDirectory() as d:
            (pathlib.Path(d) / "broken.json").write_text("{nope", encoding="utf-8")
            (pathlib.Path(d) / "partial.json").write_text(json.dumps({"bg": "#000000"}), encoding="utf-8")
            themes, problems = load_themes(d)
            self.assertEqual(themes, {})
            self.assertEqual(len(problems), 2)

    def test_mix(self):
        self.assertEqual(mix("#000000", "#ffffff", 0.5), "#808080")


if __name__ == "__main__":
    unittest.main()
