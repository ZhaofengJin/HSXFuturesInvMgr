"""
测试 main 模块
"""

import unittest
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import MagicMock, patch

import main
from models import CoilRecord, ScheduleRecord


class TestMainHelpers(unittest.TestCase):
    """测试 main 模块辅助函数"""

    @patch("main.glob.glob")
    def test_find_files_filters_temp_and_backup_and_falls_back(self, mock_glob):
        mock_glob.side_effect = [
            ["/results/~$期货库存明细.xlsx", "/results/期货库存明细_备份.xlsx"],
            [],
            ["/current/期货库存明细.xlsx"],
            ["/current/烨辉库存表.xlsx", "/current/烨辉库存表_备份.xlsx"],
        ]

        futures_files, yehui_files = main.find_files("/results", "/current")

        self.assertEqual(futures_files, ["/current/期货库存明细.xlsx"])
        self.assertEqual(yehui_files, ["/current/烨辉库存表.xlsx"])

    def test_backup_original_file_returns_false_for_missing_file(self):
        with TemporaryDirectory() as tmpdir:
            source = Path(tmpdir) / "missing.xlsx"
            backup = Path(tmpdir) / "backup.xlsx"
            self.assertFalse(main.backup_original_file(str(source), str(backup)))

    def test_setup_directories_creates_results_subdir(self):
        with TemporaryDirectory() as tmpdir:
            results_dir = main.setup_directories(tmpdir)
            self.assertTrue(Path(results_dir).is_dir())
            self.assertEqual(Path(results_dir).name, "results")


class TestMainEntryPoint(unittest.TestCase):
    """测试 main.main 主流程"""

    @patch("main.find_files", return_value=([], []))
    @patch("main.setup_directories", return_value="/tmp/results")
    @patch("main.os.chdir")
    @patch("sys.stdout", new_callable=StringIO)
    def test_main_returns_error_when_files_are_missing(self, fake_out, mock_chdir, mock_setup, mock_find):
        code = main.main("/tmp/base")
        self.assertEqual(code, 1)
        self.assertIn("Error: Excel files not found", fake_out.getvalue())

    @patch("main.format_date", return_value="2026/5/10")
    @patch("main.safe_sheet_name", return_value="客户A")
    @patch("main.InventoryProcessor")
    @patch("main.ExcelWriter")
    @patch("main.ExcelReader")
    @patch("main.load_workbook")
    @patch("main.backup_original_file", return_value=True)
    @patch("main.find_files", return_value=(["futures.xlsx"], ["yehui.xlsx"]))
    @patch("main.setup_directories", return_value="/tmp/results")
    @patch("main.os.chdir")
    def test_main_successfully_rebuilds_and_saves_workbook(
        self,
        mock_chdir,
        mock_setup,
        mock_find,
        mock_backup,
        mock_load_workbook,
        mock_reader_cls,
        mock_writer_cls,
        mock_processor_cls,
        mock_safe_sheet_name,
        mock_format_date,
    ):
        futures_wb = MagicMock()
        futures_wb.sheetnames = ["期货排程", "Sheet1", "客户A"]
        schedule_ws = MagicMock()
        futures_wb.__getitem__.side_effect = lambda name: schedule_ws if name == "期货排程" else MagicMock()
        yehui_wb = MagicMock()
        yehui_wb.sheetnames = ["烨辉"]
        yehui_wb.__getitem__.return_value = MagicMock()
        mock_load_workbook.side_effect = [futures_wb, yehui_wb]

        reader = MagicMock()
        reader.find_template_sheet.return_value = (MagicMock(title="客户A"), [(1, "倉別"), (2, "移撥日期"), (3, "入庫日期")])
        reader.read_schedule_data.return_value = (
            {"ORD001": ScheduleRecord(order_no="ORD001", customer="客户A", date="2026/4/20")},
            {},
        )
        original_row = CoilRecord(order_no="ORD001", coil_no="COIL001", customer="客户A", row_data=["v"] * 24)
        reader.read_yehui_data.return_value = ({"ORD001_COIL001": MagicMock(order_no="ORD001")}, {"COIL002": {"warehouse": "W2"}})
        reader.read_customer_data.return_value = ([original_row], [original_row])
        mock_reader_cls.return_value = reader

        writer = MagicMock()
        writer.create_customer_sheet.return_value = MagicMock()
        writer.update_schedule_colors.return_value = (1, 0)
        mock_writer_cls.return_value = writer

        processor = MagicMock()
        processor.compare_data.return_value = ([original_row], [{"order_no": "ORD001", "coil_no": "COIL001", "customer": "客户A", "row_data": ["v"] * 23}], [])
        processor.build_coil_set.return_value = {"COIL001"}
        processor.find_new_coils.return_value = [{"order_no": "ORD001", "coil_no": "COIL002", "customer": "客户A", "data": {}}]
        processor.group_by_customer.return_value = {
            "客户A": [
                ("preserved", original_row),
                ("updated", {"order_no": "ORD001", "coil_no": "COIL001", "customer": "客户A", "row_data": ["v"] * 23}),
                ("new", {"order_no": "ORD001", "coil_no": "COIL002", "customer": "客户A", "data": {}}),
            ]
        }
        processor.build_order_has_yehui.return_value = {"ORD001"}
        processor.build_order_in_customer_sheet.return_value = {"ORD001"}
        mock_processor_cls.return_value = processor

        code = main.main("/tmp/base")

        self.assertEqual(code, 0)
        writer.write_preserved_row.assert_called_once()
        writer.write_updated_row.assert_called_once()
        writer.write_new_row.assert_called_once()
        futures_wb.save.assert_called_once_with("futures.xlsx")


if __name__ == "__main__":
    unittest.main()
