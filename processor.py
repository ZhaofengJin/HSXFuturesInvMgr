"""
核心处理逻辑模块
负责数据比对、分组、更新标记等业务逻辑
"""

from collections import defaultdict
from typing import Dict, List, Tuple, Set, Any

from models import CoilRecord, ScheduleRecord, YehuiRecord


class InventoryProcessor:
    """库存数据处理器"""

    # ---------- 比对逻辑 ----------

    def compare_data(self, original_rows: List[CoilRecord],
                     yehui_coil_info: Dict[str, Dict[str, Any]],
                     schedule_data: Dict[str, ScheduleRecord]) -> Tuple[List[CoilRecord], List[Dict], List[Dict]]:
        """
        比对原文件数据与烨辉表数据
        Args:
            original_rows: 原文件中的所有钢卷记录
            yehui_coil_info: 钢卷号 -> {warehouse, transfer_date, entry_date}
            schedule_data: 订单号 -> ScheduleRecord（用于补充客户信息）
        Returns:
            (preserved_rows, updated_coils, new_coils)
        """
        preserved: List[CoilRecord] = []
        updated: List[Dict[str, Any]] = []

        for row_info in original_rows:
            coil_no = row_info.coil_no
            old_warehouse = row_info.warehouse

            if coil_no in yehui_coil_info:
                new_warehouse = yehui_coil_info[coil_no].get("warehouse", "")

                if old_warehouse != new_warehouse:
                    # 仓别有变动 → updated（橙色）
                    updated.append({
                        **row_info.to_dict(),
                        "new_warehouse": new_warehouse,
                        "new_transfer_date": yehui_coil_info[coil_no].get("transfer_date"),
                        "new_entry_date": yehui_coil_info[coil_no].get("entry_date"),
                    })
                else:
                    # 仓别无变动 → 原样保留
                    preserved.append(row_info)
            else:
                # 烨辉表中不存在该钢卷 → 仍然保留（只增加不删减）
                preserved.append(row_info)

        return preserved, updated, []

    def find_new_coils(self, schedule_data: Dict[str, ScheduleRecord],
                       yehui_data: Dict[str, YehuiRecord],
                       original_coil_set: Set[str]) -> List[Dict[str, Any]]:
        """
        以期货排程订单为基准，找出需要新增的钢卷
        Args:
            schedule_data: 订单号 -> ScheduleRecord
            yehui_data: key -> YehuiRecord
            original_coil_set: 原文件中已有的钢卷号集合
        Returns:
            新增钢卷列表
        """
        new_coils: List[Dict[str, Any]] = []

        for order_no, schedule_info in schedule_data.items():
            customer = schedule_info.customer
            if not customer:
                continue

            for key, yehui_record in yehui_data.items():
                if yehui_record.order_no != order_no:
                    continue

                coil_no = yehui_record.coil_no

                if coil_no not in original_coil_set:
                    new_coils.append({
                        "coil_no": coil_no,
                        "order_no": order_no,
                        "customer": customer,
                        "data": yehui_record.data,
                    })

        return new_coils

    # ---------- 分组逻辑 ----------

    def group_by_customer(self, preserved_rows: List[CoilRecord],
                          updated_coils: List[Dict[str, Any]],
                          new_coils: List[Dict[str, Any]],
                          schedule_data: Dict[str, ScheduleRecord]) -> Dict[str, List[Tuple[str, Any]]]:
        """
        按客户分组所有数据行
        Args:
            preserved_rows: 保留行
            updated_coils: 更新行
            new_coils: 新增行
            schedule_data: 用于补充更新行的客户信息
        Returns:
            customer_name -> [(type, data), ...]
        """
        customer_groups: Dict[str, List[Tuple[str, Any]]] = defaultdict(list)

        # 保留行
        for row_info in preserved_rows:
            customer = row_info.customer
            if customer:
                customer_groups[customer].append(("preserved", row_info))

        # 更新行
        for update_info in updated_coils:
            customer = update_info.get("customer", "")
            if not customer:
                order_no = update_info.get("order_no", "")
                customer = schedule_data.get(order_no, ScheduleRecord(order_no="", customer="")).customer
            if customer:
                customer_groups[customer].append(("updated", update_info))

        # 新增行
        for new_info in new_coils:
            customer = new_info.get("customer", "")
            if customer:
                customer_groups[customer].append(("new", new_info))

        # 排序：按订单号排序，相同类型优先 preserved
        for customer in customer_groups:
            customer_groups[customer].sort(
                key=lambda x: (x[1].get("order_no", "") if isinstance(x[1], dict) else x[1].order_no,
                               0 if x[0] == "preserved" else 1)
            )

        return dict(customer_groups)

    # ---------- 工具方法 ----------

    @staticmethod
    def build_coil_set(rows: List[CoilRecord]) -> Set[str]:
        """从行记录中提取非空钢卷号集合"""
        return {row.coil_no for row in rows if row.coil_no}

    @staticmethod
    def build_order_has_yehui(yehui_data: Dict[str, YehuiRecord]) -> Set[str]:
        """构建有烨辉数据的订单号集合"""
        return {record.order_no for record in yehui_data.values()}

    @staticmethod
    def build_order_in_customer_sheet(original_rows: List[CoilRecord]) -> Set[str]:
        """构建在客户 Sheet 中出现过的订单号集合"""
        return {row.order_no for row in original_rows if row.order_no}
