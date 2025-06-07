import unittest
from excel_mcp.utils import (
    address_within_ranges,
    collect_column_outputs,
    gather_row_context,
)


class TestAddressWithinRanges(unittest.TestCase):
    def test_address_inside(self):
        ranges = ["Sheet1!A1:C3", "Sheet2!D4:E5"]
        self.assertTrue(address_within_ranges("Sheet1!B2", ranges))

    def test_address_outside(self):
        ranges = ["Sheet1!A1:C3"]
        self.assertFalse(address_within_ranges("Sheet1!D1", ranges))

    def test_sheet_mismatch(self):
        ranges = ["Sheet1!A1:C3"]
        self.assertFalse(address_within_ranges("Sheet2!A1", ranges))


class TestGatherRowContext(unittest.TestCase):
    def test_scan_left_and_stop_right_on_text(self):
        data = {
            "A1": {"output": "a"},
            "B1": {"output": "b"},
            "C1": {"output": 3},
            "D1": {"output": 4, "formula": "=4"},
            "E1": {"output": "stop"},
        }
        result = gather_row_context(data, "C1")
        expected = {"B1": "b", "A1": "a", "C1": 3, "D1": 4, "E1": "stop"}
        self.assertEqual(result, expected)

    def test_left_limit(self):
        data = {
            "A2": {"output": "x"},
            "B2": {"output": "y"},
            "C2": {"output": "z"},
            "D2": {"output": 10},
        }
        result = gather_row_context(data, "D2")
        expected = {"C2": "z", "B2": "y", "A2": "x", "D2": 10}
        self.assertEqual(result, expected)


class TestCollectColumnOutputs(unittest.TestCase):
    def test_basic_scan(self):
        cells = {
            "A1": {"output": "Header"},
            "A2": {"output": 5},
            "A3": {"output": "x", "formula": "=B1"},
            "A4": {"output": 10},
            "A5": {"output": "stop"},
        }

        result = collect_column_outputs(cells, "A4", text_limit=1)
        expected = {
            "A3": "x",
            "A2": 5,
            "A1": "Header",
            "A4": 10,
            "A5": "stop",
        }
        self.assertEqual(result, expected)


if __name__ == "__main__":
    unittest.main()
