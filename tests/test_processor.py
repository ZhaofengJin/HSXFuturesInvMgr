"""
测试 processor 模块
TDD: 先写测试，后实现
"""

import unittest
from datetime import datetime

from processor import InventoryProcessor
from models import CoilRecord, ScheduleRecord, YehuiRecord


class TestInventoryProcessor(unittest.TestCase):
    """测试库存处理核心逻辑"""

    def setUp(self):
        self.processor = InventoryProcessor()

    # ---------- 比对逻辑测试 ----------

    def test_warehouse_changed_marked_updated(self):
        """原文件有，烨辉有，仓别变 → updated"""
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        yehui_info = {
            "COIL001": {"warehouse": "W2", "transfer_date": "2026/4/21", "entry_date": "2026/4/22"},
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(preserved), 0)
        self.assertEqual(len(updated), 1)
        self.assertEqual(len(new), 0)
        self.assertEqual(updated[0]["new_warehouse"], "W2")

    def test_warehouse_same_marked_preserved(self):
        """原文件有，烨辉有，仓别同 → preserved"""
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        yehui_info = {
            "COIL001": {"warehouse": "W1", "transfer_date": "2026/4/20", "entry_date": "2026/4/21"},
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(preserved), 1)
        self.assertEqual(len(updated), 0)
        self.assertEqual(len(new), 0)

    def test_not_in_yehui_preserved(self):
        """原文件有，烨辉无 → preserved（只增加不删减）"""
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        yehui_info = {}

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(preserved), 1)
        self.assertEqual(len(updated), 0)
        self.assertEqual(len(new), 0)

    def test_new_coil_from_schedule_and_yehui(self):
        """原文件无，期货排程有，烨辉有 → new"""
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        yehui_data = {
            "ORD001_COIL002": YehuiRecord(
                order_no="ORD001", coil_no="COIL002",
                data={"倉別": "W3", "移撥日期": "2026/4/25"},
            ),
        }
        yehui_info = {
            "COIL002": {"warehouse": "W3", "transfer_date": "2026/4/25", "entry_date": "2026/4/26"},
        }
        schedule_data = {
            "ORD001": ScheduleRecord(order_no="ORD001", customer="无锡晟明", date="2026/4/20"),
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})
        new_coils = self.processor.find_new_coils(schedule_data, yehui_data, {"COIL001"})

        self.assertEqual(len(preserved), 1)
        self.assertEqual(len(new_coils), 1)
        self.assertEqual(new_coils[0]["coil_no"], "COIL002")
        self.assertEqual(new_coils[0]["customer"], "无锡晟明")

    def test_new_coil_skipped_if_already_exists(self):
        """钢卷已存在则跳过，不重复添加"""
        yehui_data = {
            "ORD001_COIL001": YehuiRecord(
                order_no="ORD001", coil_no="COIL001",
                data={"倉別": "W1"},
            ),
        }
        schedule_data = {
            "ORD001": ScheduleRecord(order_no="ORD001", customer="无锡晟明", date="2026/4/20"),
        }
        original_coil_set = {"COIL001"}

        new_coils = self.processor.find_new_coils(schedule_data, yehui_data, original_coil_set)

        self.assertEqual(len(new_coils), 0)

    def test_empty_original_and_schedule(self):
        """空数据场景"""
        preserved, updated, new = self.processor.compare_data([], {}, {})
        self.assertEqual(len(preserved), 0)
        self.assertEqual(len(updated), 0)
        self.assertEqual(len(new), 0)

    # ---------- 分组逻辑测试 ----------

    def test_group_by_customer(self):
        """按客户分组"""
        preserved = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
            CoilRecord(order_no="ORD002", coil_no="COIL002", warehouse="W2", customer="上海客户"),
        ]
        updated = [
            {"order_no": "ORD001", "coil_no": "COIL003", "customer": "无锡晟明", "new_warehouse": "W3"},
        ]
        new = [
            {"order_no": "ORD002", "coil_no": "COIL004", "customer": "上海客户", "data": {}},
        ]
        schedule_data = {
            "ORD001": ScheduleRecord(order_no="ORD001", customer="无锡晟明"),
            "ORD002": ScheduleRecord(order_no="ORD002", customer="上海客户"),
        }

        groups = self.processor.group_by_customer(preserved, updated, new, schedule_data)

        self.assertIn("无锡晟明", groups)
        self.assertIn("上海客户", groups)
        self.assertEqual(len(groups["无锡晟明"]), 2)  # preserved + updated
        self.assertEqual(len(groups["上海客户"]), 2)  # preserved + new

    def test_group_sorting(self):
        """测试分组内按订单号排序"""
        preserved = [
            CoilRecord(order_no="ORD002", coil_no="COIL002", warehouse="W2", customer="无锡晟明"),
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]

        groups = self.processor.group_by_customer(preserved, [], [], {})
        items = groups["无锡晟明"]

        # preserved 排在前面，且按订单号排序
        self.assertEqual(items[0][1].order_no, "ORD001")
        self.assertEqual(items[1][1].order_no, "ORD002")

    # ---------- 工具方法测试 ----------

    def test_build_coil_set(self):
        """测试构建钢卷号集合"""
        rows = [
            CoilRecord(order_no="ORD001", coil_no="COIL001"),
            CoilRecord(order_no="ORD001", coil_no="COIL002"),
            CoilRecord(order_no="ORD002", coil_no=""),  # 空钢卷号应被忽略
        ]
        coil_set = self.processor.build_coil_set(rows)
        self.assertEqual(coil_set, {"COIL001", "COIL002"})

    def test_build_order_has_yehui(self):
        """测试构建订单-烨辉映射"""
        yehui_data = {
            "ORD001_COIL001": YehuiRecord(order_no="ORD001", coil_no="COIL001"),
            "ORD002_COIL002": YehuiRecord(order_no="ORD002", coil_no="COIL002"),
        }
        result = self.processor.build_order_has_yehui(yehui_data)
        self.assertEqual(result, {"ORD001", "ORD002"})

    def test_build_order_in_customer_sheet(self):
        """测试构建订单-客户Sheet映射"""
        rows = [
            CoilRecord(order_no="ORD001", coil_no="COIL001"),
            CoilRecord(order_no="ORD002", coil_no="COIL002"),
        ]
        result = self.processor.build_order_in_customer_sheet(rows)
        self.assertEqual(result, {"ORD001", "ORD002"})


if __name__ == "__main__":
    unittest.main()
