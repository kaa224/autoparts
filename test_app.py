import io
import unittest

from app import app, parse_table


class SpreadsheetValidationTest(unittest.TestCase):
    def test_columns_are_case_and_space_insensitive(self):
        rows = parse_table([
            [" Артикул ", "НАИМЕНОВАНИЕ"],
            [636116.0, " Ремкомплект   суппорта "],
        ])
        self.assertEqual(rows, [("636116", "Ремкомплект суппорта")])

    def test_missing_column_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "наименование"):
            parse_table([["артикул", "цена"], ["636116", "100"]])

    def test_non_excel_upload_is_rejected(self):
        client = app.test_client()
        response = client.post(
            "/api/jobs",
            data={"file": (io.BytesIO(bytes("артикул", "utf-8")), "parts.csv")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn(".xls", response.get_json()["error"])


if __name__ == "__main__":
    unittest.main()
