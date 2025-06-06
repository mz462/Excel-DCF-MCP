import unittest
from unittest.mock import patch
from importlib import import_module

# Import the actual module, not the server instance exposed in __init__
server_mod = import_module('excel_mcp.server')


class TestExcelEventSink(unittest.TestCase):
    def test_event_sink_records_events(self):
        sink = server_mod._ExcelEventSink()

        class DummySheet:
            def __init__(self, name):
                self.Name = name

        class DummyTarget:
            def Address(self, row_abs, col_abs):
                return "A1"

        sheet = DummySheet("Sheet1")
        target = DummyTarget()

        sink.OnSheetChange(sheet, target)
        sink.OnSheetCalculate(sheet)

        self.assertEqual(len(sink.events), 2)
        self.assertEqual(sink.events[0]["event"], "SheetChange")
        self.assertEqual(sink.events[0]["address"], "Sheet1!A1")
        self.assertEqual(sink.events[1]["event"], "SheetCalculate")


class TestServerTools(unittest.TestCase):
    def test_get_formula_no_win32(self):
        with patch.object(server_mod, "win32", None):
            result = server_mod.get_formula.fn(None, "A1")
            self.assertEqual(result["status"], "failure")

    def test_get_formula_not_initialized(self):
        with patch.object(server_mod, "win32", object()):
            with patch.object(server_mod, "excel_app", None):
                result = server_mod.get_formula.fn(None, "A1")
                self.assertEqual(result["status"], "failure")

    def test_stop_event_monitor_not_running(self):
        with patch.object(server_mod, "_event_thread", None):
            result = server_mod.stop_excel_event_monitor.fn()
            self.assertEqual(result["status"], "not_running")

    def test_fetch_events_not_running(self):
        with patch.object(server_mod, "_excel_event_handler", None):
            result = server_mod.fetch_excel_events.fn()
            self.assertEqual(result["status"], "failure")


if __name__ == "__main__":
    unittest.main()
