import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from formatter import convert_dtr, extract_employees


class FakeReader:
    def __init__(self, sheets):
        self._sheets = sheets

    def sheet(self, name):
        return self._sheets.get(name, {})


class DummyWorkbook:
    def save(self, path):
        Path(path).write_bytes(b"ok")


class FormatterTests(unittest.TestCase):
    def test_extract_employees_reads_more_than_sixteen_punch_columns(self):
        reader = FakeReader(
            {
                "Att. Stat.": {
                    (5, 1): "1001",
                    (5, 2): "Juan Dela Cruz",
                    (5, 3): "IT",
                },
                "Att.log report": {
                    (1, 1): "ID:",
                    (1, 3): "1001",
                    (2, 1): "09:00",
                    (2, 2): "17:00",
                    (2, 3): "09:00",
                    (2, 4): "17:00",
                    (2, 5): "09:00",
                    (2, 6): "17:00",
                    (2, 7): "09:00",
                    (2, 8): "17:00",
                    (2, 9): "09:00",
                    (2, 10): "17:00",
                    (2, 11): "09:00",
                    (2, 12): "17:00",
                    (2, 13): "09:00",
                    (2, 14): "17:00",
                    (2, 15): "09:00",
                    (2, 16): "17:00",
                    (2, 17): "09:00",
                    (2, 18): "17:00",
                    (2, 19): "09:00",
                    (2, 20): "17:00",
                    (2, 21): "09:00",
                    (2, 22): "17:00",
                    (2, 23): "09:00",
                    (2, 24): "17:00",
                    (2, 25): "09:00",
                    (2, 26): "17:00",
                    (2, 27): "09:00",
                    (2, 28): "17:00",
                    (2, 29): "09:00",
                    (2, 30): "17:00",
                },
            }
        )

        employees = extract_employees(reader)

        self.assertEqual(1, len(employees))
        self.assertEqual(30, len(employees[0].punches or []))
        self.assertEqual("09:00", employees[0].punches[0])
        self.assertEqual("17:00", employees[0].punches[-1])

    def test_convert_dtr_preserves_full_period(self):
        class FakeXlsReader:
            def __init__(self, data):
                self.data = data

            def sheet(self, name):
                if name == "Att. Stat.":
                    return {
                        (5, 1): "1001",
                        (5, 2): "Juan Dela Cruz",
                        (5, 3): "IT",
                    }
                if name == "Att.log report":
                    return {
                        (1, 1): "ID:",
                        (1, 3): "1001",
                        (2, 1): "09:00",
                        (2, 2): "17:00",
                    }
                if name in {"Exception Stat.", "Schedule Infor."}:
                    return { (1, 1): "2024-01-01 ~ 2024-01-30" }
                return {}

        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "sample.xls"
            input_path.write_bytes(b"not-a-real-xls")
            with patch("formatter.XlsReader", FakeXlsReader), patch("formatter.build_output", return_value=DummyWorkbook()):
                output_path, info = convert_dtr(input_path, "sample.xls", Path(tmp_dir))

            self.assertTrue(output_path.exists())
            self.assertEqual(30, info["days"])


if __name__ == "__main__":
    unittest.main()
