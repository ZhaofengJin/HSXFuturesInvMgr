"""
测试 utils 模块
TDD: 先写测试，后实现
"""

import unittest
from datetime import datetime

from utils import format_date, get_field_by_name, safe_sheet_name, find_base_dir
from config import STANDARD_HEADERS, FULL_HEADERS


class TestFormatDate(unittest.TestCase):
    """测试日期格式化"""

    def test_none_returns_empty(self):
        self.assertEqual(format_date(None), "")

    def test_datetime_format(self):
        dt = datetime(2026, 4, 20)
        self.assertEqual(format_date(dt), "2026/4/20")

    def test_datetime_with_time(self):
        dt = datetime(2026, 4, 20, 10, 30, 0)
        self.assertEqual(format_date(dt), "2026/4/20")

    def test_iso_string(self):
        self.assertEqual(format_date("2026-04-20"), "2026/4/20")

    def test_iso_string_with_time(self):
        self.assertEqual(format_date("2026-04-20 10:30:00"), "2026/4/20")

    def test_already_formatted_string(self):
        self.assertEqual(format_date("2026/4/20"), "2026/4/20")

    def test_invalid_string_returns_original(self):
        self.assertEqual(format_date("not a date"), "not a date")

    def test_number_returns_str(self):
        self.assertEqual(format_date(2026), "2026")


class TestGetFieldByName(unittest.TestCase):
    """测试字段名变体匹配"""

    def test_exact_match(self):
        row = {"倉別": "A仓库"}
        self.assertEqual(get_field_by_name(row, "倉別"), "A仓库")

    def test_variant_match_simplified(self):
        row = {"仓别": "B仓库"}
        self.assertEqual(get_field_by_name(row, "倉別"), "B仓库")

    def test_variant_match_english(self):
        row = {"Warehouse": "C仓库"}
        self.assertEqual(get_field_by_name(row, "倉別"), "C仓库")

    def test_no_match_returns_none(self):
        row = {"其他字段": "value"}
        self.assertIsNone(get_field_by_name(row, "倉別"))

    def test_multiple_variants_priority(self):
        # 优先返回第一个匹配的
        row = {"倉別": "A", "仓别": "B"}
        self.assertEqual(get_field_by_name(row, "倉別"), "A")

    def test_empty_row(self):
        self.assertIsNone(get_field_by_name({}, "倉別"))

    def test_all_key_fields(self):
        """测试所有关键字段的变体匹配"""
        row = {
            "订单编号": "ORD001",
            "钢卷号": "COIL001",
            "仓库": "W1",
            "调拨日期": "2026/4/20",
            "入库日期": "2026/4/21",
        }
        self.assertEqual(get_field_by_name(row, "訂單編號"), "ORD001")
        self.assertEqual(get_field_by_name(row, "鋼捲編號"), "COIL001")
        self.assertEqual(get_field_by_name(row, "倉別"), "W1")
        self.assertEqual(get_field_by_name(row, "移撥日期"), "2026/4/20")
        self.assertEqual(get_field_by_name(row, "入庫日期"), "2026/4/21")


class TestSafeSheetName(unittest.TestCase):
    """测试安全 Sheet 名转换"""

    def test_normal_name(self):
        self.assertEqual(safe_sheet_name("无锡晟明"), "无锡晟明")

    def test_truncate_long_name(self):
        long_name = "A" * 40
        self.assertEqual(len(safe_sheet_name(long_name)), 31)

    def test_replace_invalid_chars(self):
        self.assertEqual(
            safe_sheet_name("客户/名称[测试]"),
            "客户_名称_测试_"
        )

    def test_mixed_invalid_and_long(self):
        name = "A" * 20 + "/" + "B" * 20
        result = safe_sheet_name(name)
        self.assertNotIn("/", result)
        self.assertLessEqual(len(result), 31)


class TestFindBaseDir(unittest.TestCase):
    """测试查找基础目录"""

    def test_find_existing_keyword_dir(self):
        import os
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            target = os.path.join(tmpdir, "期货库存管理")
            os.makedirs(target)
            found = find_base_dir(tmpdir, "期货")
            self.assertEqual(found, target)

    def test_no_match_returns_none(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            found = find_base_dir(tmpdir, "不存在的关键字")
            self.assertIsNone(found)

    def test_skip_temp_files(self):
        import os
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            # 创建 ~$ 开头的目录（Excel 临时文件模式）
            temp_dir = os.path.join(tmpdir, "~$期货")
            os.makedirs(temp_dir)
            found = find_base_dir(tmpdir, "期货")
            self.assertIsNone(found)


class TestConfigConstants(unittest.TestCase):
    """测试配置常量"""

    def test_standard_headers_count(self):
        self.assertEqual(len(STANDARD_HEADERS), 23)

    def test_full_headers_count(self):
        self.assertEqual(len(FULL_HEADERS), 24)

    def test_modify_date_last(self):
        self.assertEqual(FULL_HEADERS[-1], "修改日期")


if __name__ == "__main__":
    unittest.main()
