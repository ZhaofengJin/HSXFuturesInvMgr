"""
测试 models 模块
TDD: 先写测试，后实现
"""

import unittest
from datetime import datetime

from models import CoilRecord, ScheduleRecord, YehuiRecord, ProcessingResult


class TestCoilRecord(unittest.TestCase):
    """测试钢卷记录模型"""

    def test_basic_creation(self):
        record = CoilRecord(
            order_no="ORD001",
            coil_no="COIL001",
            warehouse="W1",
            transfer_date="2026/4/20",
            entry_date="2026/4/21",
            customer="无锡晟明",
            row_data=["2026/4/20", "无锡晟明", "ORD001", "COIL001"],
        )
        self.assertEqual(record.order_no, "ORD001")
        self.assertEqual(record.coil_no, "COIL001")
        self.assertEqual(record.warehouse, "W1")

    def test_default_values(self):
        record = CoilRecord(order_no="ORD001", coil_no="COIL001")
        self.assertEqual(record.warehouse, "")
        self.assertIsNone(record.transfer_date)
        self.assertIsNone(record.entry_date)
        self.assertEqual(record.customer, "")
        self.assertEqual(record.row_data, [])
        self.assertIsNone(record.modify_date)

    def test_to_dict(self):
        record = CoilRecord(
            order_no="ORD001",
            coil_no="COIL001",
            warehouse="W1",
            transfer_date="2026/4/20",
            entry_date="2026/4/21",
            customer="无锡晟明",
        )
        d = record.to_dict()
        self.assertEqual(d["order_no"], "ORD001")
        self.assertEqual(d["coil_no"], "COIL001")
        self.assertEqual(d["warehouse"], "W1")

    def test_to_dict_includes_fill(self):
        """to_dict 必须包含 fill，否则保留行颜色会在 cli.py 传递时丢失"""
        from openpyxl.styles import PatternFill
        orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        record = CoilRecord(
            order_no="ORD001",
            coil_no="COIL001",
            warehouse="W1",
            fill=orange_fill,
            modify_date="2026/4/15",
        )
        d = record.to_dict()
        self.assertEqual(d["fill"], orange_fill)
        self.assertEqual(d["modify_date"], "2026/4/15")


class TestScheduleRecord(unittest.TestCase):
    """测试期货排程记录模型"""

    def test_basic_creation(self):
        record = ScheduleRecord(
            order_no="ORD001",
            customer="无锡晟明",
            date="2026/4/20",
            row=2,
        )
        self.assertEqual(record.order_no, "ORD001")
        self.assertEqual(record.customer, "无锡晟明")
        self.assertEqual(record.date, "2026/4/20")
        self.assertEqual(record.row, 2)


class TestYehuiRecord(unittest.TestCase):
    """测试烨辉记录模型"""

    def test_basic_creation(self):
        record = YehuiRecord(
            order_no="ORD001",
            coil_no="COIL001",
            data={"倉別": "W1", "移撥日期": "2026/4/20"},
        )
        self.assertEqual(record.order_no, "ORD001")
        self.assertEqual(record.coil_no, "COIL001")
        self.assertEqual(record.data["倉別"], "W1")

    def test_warehouse_property(self):
        record = YehuiRecord(
            order_no="ORD001",
            coil_no="COIL001",
            data={"倉別": "W1", "移撥日期": "2026/4/20", "入庫日期": "2026/4/21"},
        )
        self.assertEqual(record.warehouse, "W1")
        self.assertEqual(record.transfer_date, "2026/4/20")
        self.assertEqual(record.entry_date, "2026/4/21")

    def test_warehouse_property_missing(self):
        record = YehuiRecord(
            order_no="ORD001",
            coil_no="COIL001",
            data={},
        )
        self.assertEqual(record.warehouse, "")
        self.assertIsNone(record.transfer_date)
        self.assertIsNone(record.entry_date)


class TestProcessingResult(unittest.TestCase):
    """测试处理结果模型"""

    def test_empty_result(self):
        result = ProcessingResult()
        self.assertEqual(len(result.preserved), 0)
        self.assertEqual(len(result.updated), 0)
        self.assertEqual(len(result.new), 0)
        self.assertEqual(result.matched_count, 0)
        self.assertEqual(result.unmatched_count, 0)

    def test_summary(self):
        result = ProcessingResult()
        result.preserved = [1, 2]
        result.updated = [3]
        result.new = [4, 5, 6]
        result.matched_count = 10
        result.unmatched_count = 2

        summary = result.summary()
        self.assertEqual(summary["preserved"], 2)
        self.assertEqual(summary["updated"], 1)
        self.assertEqual(summary["new"], 3)
        self.assertEqual(summary["matched"], 10)
        self.assertEqual(summary["unmatched"], 2)


if __name__ == "__main__":
    unittest.main()
