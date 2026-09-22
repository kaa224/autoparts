import unittest

from PIL import Image, ImageDraw

from generate_card import (
    Application, applicability_labels, band_lines, fit_wrapped_text,
    normalize_article, split_balanced,
)


class GeneratorRulesTest(unittest.TestCase):
    def test_numeric_xls_article(self):
        self.assertEqual(normalize_article(636116.0), "636116")

    def test_six_makes_outputs_only_makes(self):
        apps = tuple(Application(make, f"{make} MODEL 2000-2001", ()) for make in "A B C D E F".split())
        self.assertEqual(applicability_labels(apps), list("ABCDEF"))

    def test_three_makes_outputs_make_and_model(self):
        apps = (
            Application("FORD", "FORD FOCUS 2000-2004", ("FORD FOCUS 2.0L 2000-2004",)),
            Application("MAZDA", "MAZDA 3 2004-2008", ("MAZDA 3 2.0L 2004-2008",)),
            Application("VOLVO", "VOLVO S40 2004-2012", ("VOLVO S40 2.4L 2004-2012",)),
        )
        self.assertEqual(applicability_labels(apps), ["FORD FOCUS", "MAZDA 3", "VOLVO S40"])

    def test_two_makes_outputs_details(self):
        apps = (
            Application("FORD", "FORD FOCUS 2000-2004", ("FORD FOCUS 2.0L 2000-2004",)),
            Application("MAZDA", "MAZDA 3 2004-2008", ("MAZDA 3 2.0L 2004-2008",)),
        )
        self.assertEqual(applicability_labels(apps), ["FORD FOCUS 2.0L 2000-2004", "MAZDA 3 2.0L 2004-2008"])

    def test_six_make_model_combinations_reduce_to_makes(self):
        apps = tuple(
            Application(make, f"{make} MODEL{i} 2000-2004", ())
            for make, i in (("FORD", 1), ("FORD", 2), ("MAZDA", 1), ("MAZDA", 2), ("VOLVO", 1), ("VOLVO", 2))
        )
        self.assertEqual(applicability_labels(apps), ["FORD", "MAZDA", "VOLVO"])

    def test_long_band_uses_two_lines(self):
        labels = ["LAND ROVER", "MERCEDES BENZ", "ALFA ROMEO", "ROLLS ROYCE"]
        lines = band_lines(labels, character_limit=30)
        self.assertEqual(len(lines), 2)
        self.assertEqual(" • ".join(lines), " • ".join(labels))

    def test_long_name_is_fully_wrapped(self):
        draw = ImageDraw.Draw(Image.new("RGB", (750, 1000)))
        text = "КОМПЛЕКТ ИЗ ПОРШНЯ ПЫЛЬНИКОВ ВТУЛОК УПЛОТНИТЕЛЬНЫХ КОЛЕЦ САЛЬНИКА И ШТУЦЕРОВ"
        _, lines, _ = fit_wrapped_text(draw, text, 690, 94)
        self.assertEqual(" ".join(lines), text)

    def test_balanced_split_keeps_all_values(self):
        top, bottom = split_balanced(["AA", "BBBB", "CCC"])
        self.assertEqual(f"{top} • {bottom}", "AA • BBBB • CCC")


if __name__ == "__main__":
    unittest.main()
