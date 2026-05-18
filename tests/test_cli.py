"""
测试 cli 模块
TDD: 先写测试，后实现
"""

import unittest
from unittest.mock import patch, MagicMock
import json
import sys
from io import StringIO

from cli import parse_args, main, format_summary_json, log


class TestParseArgs(unittest.TestCase):
    """测试参数解析"""

    def test_default_args(self):
        args = parse_args([])
        self.assertIsNone(args.base_dir)
        self.assertFalse(args.json)
        self.assertFalse(args.dry_run)
        self.assertFalse(args.verbose)

    def test_base_dir(self):
        args = parse_args(["--base-dir", "C:\\test"])
        self.assertEqual(args.base_dir, "C:\\test")

    def test_short_base_dir(self):
        args = parse_args(["-b", "C:\\test"])
        self.assertEqual(args.base_dir, "C:\\test")

    def test_json_flag(self):
        args = parse_args(["--json"])
        self.assertTrue(args.json)

    def test_dry_run_flag(self):
        args = parse_args(["--dry-run"])
        self.assertTrue(args.dry_run)

    def test_verbose_flag(self):
        args = parse_args(["-v"])
        self.assertTrue(args.verbose)

    def test_version_flag(self):
        args = parse_args(["--version"])
        self.assertTrue(args.version)


class TestFormatSummaryJson(unittest.TestCase):
    """测试 JSON 摘要格式化"""

    def test_basic_summary(self):
        summary = {
            "preserved": 10,
            "updated": 2,
            "new": 3,
            "matched": 15,
            "unmatched": 1,
        }
        result = format_summary_json(summary, exit_code=0)
        data = json.loads(result)
        self.assertEqual(data["status"], "success")
        self.assertEqual(data["data"]["preserved"], 10)
        self.assertEqual(data["data"]["updated"], 2)
        self.assertEqual(data["data"]["new"], 3)
        self.assertEqual(data["exit_code"], 0)
        self.assertIn("timestamp", data)

    def test_error_summary(self):
        result = format_summary_json(None, exit_code=1, error_msg="File not found")
        data = json.loads(result)
        self.assertEqual(data["status"], "error")
        self.assertEqual(data["error"], "File not found")
        self.assertEqual(data["exit_code"], 1)

    def test_warning_summary(self):
        result = format_summary_json({"unmatched": 1}, exit_code=2)
        data = json.loads(result)
        self.assertEqual(data["status"], "warning")
        self.assertEqual(data["exit_code"], 2)


class TestLog(unittest.TestCase):
    """测试日志输出辅助函数"""

    def test_json_mode_suppresses_regular_logs(self):
        with patch("sys.stdout", new=StringIO()) as fake_out:
            log("hidden", json_mode=True, verbose=False)
        self.assertEqual(fake_out.getvalue(), "")

    def test_verbose_log_requires_verbose_flag(self):
        with patch("sys.stdout", new=StringIO()) as fake_out:
            log("hidden", json_mode=False, verbose=False, is_verbose=True)
        self.assertEqual(fake_out.getvalue(), "")

    def test_log_prints_when_allowed(self):
        with patch("sys.stdout", new=StringIO()) as fake_out:
            log("shown", json_mode=False, verbose=True, is_verbose=True)
        self.assertIn("shown", fake_out.getvalue())


class TestCliMain(unittest.TestCase):
    """测试 CLI 主函数"""

    def test_version_output(self):
        """测试版本信息输出"""
        with patch("sys.stdout", new=StringIO()) as fake_out:
            code = main(["--version"])
        self.assertEqual(code, 0)
        self.assertIn("v5.0", fake_out.getvalue())

    @patch("os.chdir")
    @patch("cli.InventoryProcessor")
    @patch("cli.ExcelReader")
    @patch("cli.ExcelWriter")
    @patch("openpyxl.load_workbook")
    @patch("cli.find_files")
    @patch("cli.setup_directories")
    @patch("cli.backup_original_file")
    def test_dry_run_does_not_save(self, mock_backup, mock_setup, mock_find,
                                    mock_load_wb, mock_writer, mock_reader, mock_processor, mock_chdir):
        """测试 dry-run 模式不保存文件"""
        # Mock 文件查找
        mock_find.return_value = (["futures.xlsx"], ["yehui.xlsx"])
        
        # Mock workbook
        mock_wb = MagicMock()
        mock_wb.sheetnames = ["期货排程", "无锡晟明"]
        mock_load_wb.return_value = mock_wb
        
        # Mock processor
        processor = MagicMock()
        processor.compare_data.return_value = ([], [], [])
        processor.find_new_coils.return_value = []
        processor.build_coil_set.return_value = set()
        processor.build_order_has_yehui.return_value = set()
        processor.build_order_in_customer_sheet.return_value = set()
        processor.group_by_customer.return_value = {}
        mock_processor.return_value = processor
        
        # Mock reader
        reader = MagicMock()
        reader.find_template_sheet.return_value = (MagicMock(), [])
        reader.read_schedule_data.return_value = ({}, {})
        reader.read_yehui_data.return_value = ({}, {})
        reader.read_customer_data.return_value = ([], [])
        mock_reader.return_value = reader
        
        with patch("sys.stdout", new=StringIO()):
            code = main(["--dry-run", "--json", "--base-dir", "/tmp/fake_dir"])
        
        self.assertEqual(code, 0)
        # dry-run 不保存
        mock_wb.save.assert_not_called()

    @patch("cli.find_files")
    @patch("os.chdir")
    def test_fallback_to_script_dir(self, mock_chdir, mock_find_files):
        """测试无参数时默认使用脚本所在目录"""
        mock_find_files.return_value = ([], [])
        with patch("sys.stdout", new=StringIO()) as fake_out:
            code = main([])
        self.assertEqual(code, 1)
        output = fake_out.getvalue()
        # 默认使用脚本目录后，因找不到文件而报错
        self.assertIn("Excel files not found", output)

    @patch("cli.find_files")
    @patch("os.chdir")
    def test_missing_excel_files(self, mock_chdir, mock_find_files):
        """测试 Excel 文件不存在时的错误处理"""
        mock_find_files.return_value = ([], [])
        with patch("sys.stdout", new=StringIO()) as fake_out:
            code = main([])
        self.assertEqual(code, 1)
        output = fake_out.getvalue()
        self.assertIn("Excel files not found", output)

    @patch("cli.setup_directories")
    @patch("cli.find_files")
    @patch("os.chdir")
    @patch("openpyxl.load_workbook", side_effect=Exception("locked"))
    def test_load_workbook_failure_outputs_json_error(
        self, mock_load_wb, mock_chdir, mock_find_files, mock_setup
    ):
        mock_setup.return_value = "/tmp/results"
        mock_find_files.return_value = (["futures.xlsx"], ["yehui.xlsx"])
        with patch("sys.stdout", new=StringIO()) as fake_out:
            code = main(["--json"])
        self.assertEqual(code, 1)
        data = json.loads(fake_out.getvalue())
        self.assertEqual(data["status"], "error")
        self.assertIn("Failed to load Excel files", data["error"])

    @patch("os.chdir")
    @patch("cli.InventoryProcessor")
    @patch("cli.ExcelReader")
    @patch("cli.ExcelWriter")
    @patch("openpyxl.load_workbook")
    @patch("cli.find_files")
    @patch("cli.setup_directories")
    def test_missing_template_sheet_outputs_json_error(
        self, mock_setup, mock_find, mock_load_wb, mock_writer, mock_reader, mock_processor, mock_chdir
    ):
        mock_setup.return_value = "/tmp/results"
        mock_find.return_value = (["futures.xlsx"], ["yehui.xlsx"])
        mock_load_wb.return_value = MagicMock()
        mock_reader.return_value.find_template_sheet.return_value = (None, [])

        with patch("sys.stdout", new=StringIO()) as fake_out:
            code = main(["--json"])

        self.assertEqual(code, 1)
        data = json.loads(fake_out.getvalue())
        self.assertEqual(data["status"], "error")
        self.assertIn("No valid template sheet found", data["error"])

    @patch("os.chdir")
    @patch("cli.InventoryProcessor")
    @patch("cli.ExcelReader")
    @patch("cli.ExcelWriter")
    @patch("openpyxl.load_workbook")
    @patch("cli.find_files")
    @patch("cli.setup_directories")
    def test_dry_run_returns_warning_exit_code_for_unmatched_orders(
        self, mock_setup, mock_find, mock_load_wb, mock_writer, mock_reader, mock_processor, mock_chdir
    ):
        mock_setup.return_value = "/tmp/results"
        mock_find.return_value = (["futures.xlsx"], ["yehui.xlsx"])

        futures_wb = MagicMock()
        futures_wb.sheetnames = ["期货排程"]
        yehui_wb = MagicMock()
        yehui_wb.sheetnames = ["烨辉"]
        yehui_wb.__getitem__.return_value = MagicMock()
        mock_load_wb.side_effect = [futures_wb, yehui_wb]

        reader = MagicMock()
        reader.find_template_sheet.return_value = (MagicMock(title="模板"), [(1, "倉別"), (2, "移撥日期"), (3, "入庫日期")])
        reader.read_schedule_data.return_value = (
            {
                "ORD001": MagicMock(customer="客户A"),
                "ORD002": MagicMock(customer="客户B"),
            },
            {},
        )
        reader.read_yehui_data.return_value = ({}, {})
        mock_reader.return_value = reader

        processor = MagicMock()
        processor.compare_data.return_value = ([], [], [])
        processor.build_coil_set.return_value = set()
        processor.find_new_coils.return_value = []
        processor.build_order_has_yehui.return_value = {"ORD001"}
        processor.build_order_in_customer_sheet.return_value = set()
        mock_processor.return_value = processor

        with patch("sys.stdout", new=StringIO()) as fake_out:
            code = main(["--dry-run", "--json", "--base-dir", "/tmp/fake_dir"])

        self.assertEqual(code, 2)
        data = json.loads(fake_out.getvalue())
        self.assertEqual(data["status"], "warning")
        self.assertEqual(data["data"]["matched"], 1)
        self.assertEqual(data["data"]["unmatched"], 1)


class TestExitCodes(unittest.TestCase):
    """测试退出码规范"""

    def test_success_code(self):
        """成功返回 0"""
        result = format_summary_json({"preserved": 1}, 0)
        data = json.loads(result)
        self.assertEqual(data["status"], "success")
        self.assertEqual(data["exit_code"], 0)
        self.assertEqual(data["data"]["preserved"], 1)

    def test_error_code(self):
        """错误返回 1"""
        result = format_summary_json(None, 1, "error")
        data = json.loads(result)
        self.assertEqual(data["exit_code"], 1)


if __name__ == "__main__":
    unittest.main()
