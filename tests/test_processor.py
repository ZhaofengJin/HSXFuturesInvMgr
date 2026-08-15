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

    def test_warehouse_same_marked_preserved(self):
        """原文件有，烨辉有，仓别同且其他属性同 → preserved"""
        row_data = [""] * 24
        row_data[2] = "ORD001"
        row_data[3] = "COIL001"
        row_data[11] = "W1"
        row_data[12] = "2026/4/20"
        row_data[13] = "2026/4/21"
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_data=row_data),
        ]
        yehui_info = {
            "COIL001": {
                "order_no": "ORD001",
                "warehouse": "W1",
                "transfer_date": "2026/4/20",
                "entry_date": "2026/4/21",
                "data": {"訂單編號": "ORD001", "倉別": "W1", "移撥日期": "2026/4/20", "入庫日期": "2026/4/21"},
            },
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(preserved), 1)
        self.assertEqual(len(updated), 0)
        self.assertEqual(len(new), 0)

    def test_warehouse_changed_marked_updated(self):
        """原文件有，烨辉有，仓别变 → updated"""
        row_data = [""] * 24
        row_data[2] = "ORD001"
        row_data[3] = "COIL001"
        row_data[11] = "W1"
        row_data[12] = "2026/4/20"
        row_data[13] = "2026/4/21"
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_data=row_data),
        ]
        yehui_info = {
            "COIL001": {
                "order_no": "ORD001",
                "warehouse": "W2",
                "transfer_date": "2026/4/20",
                "entry_date": "2026/4/21",
                "data": {"訂單編號": "ORD001", "倉別": "W2", "移撥日期": "2026/4/20", "入庫日期": "2026/4/21"},
            },
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(preserved), 0)
        self.assertEqual(len(updated), 1)
        self.assertEqual(len(new), 0)
        self.assertEqual(updated[0]["change_type"], "updated")
        self.assertEqual(updated[0]["yehui_data"].get("倉別"), "W2")

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
        row_data = [""] * 24
        row_data[2] = "ORD001"
        row_data[3] = "COIL001"
        row_data[11] = "W1"
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_data=row_data),
        ]
        yehui_data = {
            "ORD001_COIL001": YehuiRecord(
                order_no="ORD001", coil_no="COIL001",
                data={"訂單編號": "ORD001", "倉別": "W1", "移撥日期": "2026/4/20", "入庫日期": "2026/4/21"},
            ),
            "ORD001_COIL002": YehuiRecord(
                order_no="ORD001", coil_no="COIL002",
                data={"訂單編號": "ORD001", "倉別": "W3", "移撥日期": "2026/4/25"},
            ),
        }
        yehui_info = {
            "COIL001": {"order_no": "ORD001", "warehouse": "W1", "data": {"訂單編號": "ORD001", "倉別": "W1"}},
            "COIL002": {"order_no": "ORD001", "warehouse": "W3", "data": {"訂單編號": "ORD001", "倉別": "W3"}},
        }
        schedule_data = {
            "ORD001": ScheduleRecord(order_no="ORD001", customer="无锡晟明", date="2026/4/20"),
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, schedule_data)
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

    def test_mixed_rows_keep_original_order(self):
        """同一订单号内，preserved/updated 按原文件行号混合排列，保持原有顺序"""
        preserved = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_idx=2),
            CoilRecord(order_no="ORD001", coil_no="COIL003", warehouse="W2", customer="无锡晟明", row_idx=4),
        ]
        updated = [
            {"order_no": "ORD001", "coil_no": "COIL002", "customer": "无锡晟明", "new_warehouse": "W3", "row_idx": 3},
            {"order_no": "ORD001", "coil_no": "COIL004", "customer": "无锡晟明", "new_warehouse": "W4", "row_idx": 5},
        ]

        groups = self.processor.group_by_customer(preserved, updated, [], {})
        items = groups["无锡晟明"]

        # 验证顺序：COIL001(row2,preserved), COIL002(row3,updated), COIL003(row4,preserved), COIL004(row5,updated)
        coil_nos = [item[1].coil_no if hasattr(item[1], 'coil_no') else item[1].get('coil_no') for item in items]
        self.assertEqual(coil_nos, ["COIL001", "COIL002", "COIL003", "COIL004"])

    def test_new_rows_appended_after_order_data(self):
        """新增行追加在相同订单号已有数据的下方"""
        preserved = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_idx=2),
        ]
        updated = [
            {"order_no": "ORD001", "coil_no": "COIL002", "customer": "无锡晟明", "new_warehouse": "W3", "row_idx": 3},
        ]
        new = [
            {"order_no": "ORD001", "coil_no": "COIL005", "customer": "无锡晟明", "data": {}},
            {"order_no": "ORD001", "coil_no": "COIL006", "customer": "无锡晟明", "data": {}},
        ]

        groups = self.processor.group_by_customer(preserved, updated, new, {})
        items = groups["无锡晟明"]

        # 验证顺序：preserved(row2), updated(row3), new, new
        types = [item[0] for item in items]
        self.assertEqual(types, ["preserved", "updated", "new", "new"])

        # 验证钢卷号顺序
        coil_nos = [item[1].coil_no if hasattr(item[1], 'coil_no') else item[1].get('coil_no') for item in items]
        self.assertEqual(coil_nos, ["COIL001", "COIL002", "COIL005", "COIL006"])

    def test_new_rows_appended_only_when_same_order_no(self):
        """新增行只追加在相同订单号下方，不同订单号按原有顺序排列"""
        preserved = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_idx=2),
            CoilRecord(order_no="ORD002", coil_no="COIL002", warehouse="W2", customer="无锡晟明", row_idx=3),
        ]
        new = [
            {"order_no": "ORD001", "coil_no": "COIL003", "customer": "无锡晟明", "data": {}},
            {"order_no": "ORD002", "coil_no": "COIL004", "customer": "无锡晟明", "data": {}},
        ]

        groups = self.processor.group_by_customer(preserved, [], new, {})
        items = groups["无锡晟明"]

        # ORD001 的数据在一起，ORD002 的数据在一起
        order_nos = [item[1].order_no if hasattr(item[1], 'order_no') else item[1].get('order_no') for item in items]
        self.assertEqual(order_nos, ["ORD001", "ORD001", "ORD002", "ORD002"])

        # 每个订单号内：preserved 在前，new 在后
        self.assertEqual(items[0][0], "preserved")  # ORD001
        self.assertEqual(items[1][0], "new")        # ORD001
        self.assertEqual(items[2][0], "preserved")  # ORD002
        self.assertEqual(items[3][0], "new")        # ORD002

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

    # ---------- 转单逻辑测试 ----------

    def test_order_changed_marked_transferred(self):
        """原文件有，烨辉有，订单号变化 → transferred（紫色）"""
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        yehui_info = {
            "COIL001": {
                "order_no": "ORD002",
                "warehouse": "W1",
                "transfer_date": "2026/4/20",
                "entry_date": "2026/4/21",
                "data": {"訂單編號": "ORD002", "倉別": "W1", "移撥日期": "2026/4/20", "入庫日期": "2026/4/21"},
            },
        }
        schedule_data = {
            "ORD002": ScheduleRecord(order_no="ORD002", customer="上海客户", date="2026/4/20"),
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, schedule_data)

        self.assertEqual(len(preserved), 0)
        self.assertEqual(len(updated), 1)
        self.assertEqual(len(new), 0)
        self.assertEqual(updated[0]["change_type"], "transferred")
        self.assertEqual(updated[0]["order_no"], "ORD002")
        self.assertEqual(updated[0]["new_order_no"], "ORD002")
        self.assertEqual(updated[0]["old_order_no"], "ORD001")

    def test_transferred_takes_priority_over_attribute_change(self):
        """转单+属性变更同时发生时，优先级为 transferred"""
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        yehui_info = {
            "COIL001": {
                "order_no": "ORD002",
                "warehouse": "W2",
                "transfer_date": "2026/4/20",
                "entry_date": "2026/4/21",
                "data": {"訂單編號": "ORD002", "倉別": "W2", "移撥日期": "2026/4/20", "入庫日期": "2026/4/21"},
            },
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(updated), 1)
        self.assertEqual(updated[0]["change_type"], "transferred")

    def test_date_normalization_avoids_false_positive(self):
        """日期格式不同但实质相同 → preserved"""
        from datetime import datetime
        row_data = [""] * 24
        row_data[2] = "ORD001"
        row_data[3] = "COIL001"
        row_data[11] = "W1"
        row_data[13] = "2026/8/3"
        original = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明", row_data=row_data),
        ]
        yehui_info = {
            "COIL001": {
                "order_no": "ORD001",
                "warehouse": "W1",
                "transfer_date": None,
                "entry_date": datetime(2026, 8, 3),
                "data": {"訂單編號": "ORD001", "倉別": "W1", "移撥日期": None, "入庫日期": datetime(2026, 8, 3)},
            },
        }

        preserved, updated, new = self.processor.compare_data(original, yehui_info, {})

        self.assertEqual(len(preserved), 1)
        self.assertEqual(len(updated), 0)

    def test_group_by_customer_transferred_uses_new_customer(self):
        """转单行按新订单号从排程取客户归属"""
        preserved = [
            CoilRecord(order_no="ORD001", coil_no="COIL001", warehouse="W1", customer="无锡晟明"),
        ]
        updated = [
            {
                "order_no": "ORD002",
                "coil_no": "COIL001",
                "customer": "无锡晟明",
                "change_type": "transferred",
                "new_order_no": "ORD002",
                "old_order_no": "ORD001",
            },
        ]
        schedule_data = {
            "ORD002": ScheduleRecord(order_no="ORD002", customer="上海客户（转单）"),
        }

        groups = self.processor.group_by_customer(preserved, updated, [], schedule_data)

        self.assertIn("上海客户（转单）", groups)
        self.assertEqual(len(groups["上海客户（转单）"]), 1)
        self.assertEqual(groups["上海客户（转单）"][0][0], "transferred")
        # 保留行仍留在原客户 sheet
        self.assertIn("无锡晟明", groups)
        self.assertEqual(len(groups["无锡晟明"]), 1)
        self.assertEqual(groups["无锡晟明"][0][0], "preserved")

    def test_build_order_set_from_groups(self):
        """从最终分组中构建订单号集合"""
        groups = {
            "无锡晟明": [("preserved", CoilRecord(order_no="ORD001", coil_no="COIL001")),
                        ("updated", {"order_no": "ORD002", "coil_no": "COIL002"})],
        }
        orders = self.processor.build_order_set_from_groups(groups)
        self.assertEqual(orders, {"ORD001", "ORD002"})


if __name__ == "__main__":
    unittest.main()
