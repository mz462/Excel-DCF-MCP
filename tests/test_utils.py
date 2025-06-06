import unittest
from excel_mcp.utils import address_within_ranges


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


if __name__ == "__main__":
    unittest.main()
