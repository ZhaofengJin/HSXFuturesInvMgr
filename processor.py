"""
核心处理逻辑模块
负责数据比对、分组、更新标记等业务逻辑
"""

from collections import defaultdict
from typing import Dict, List, Tuple, Set, Any

from config import STANDARD_HEADERS, FULL_HEADERS
from models import CoilRecord, ScheduleRecord, YehuiRecord
from utils import format_date, get_field_by_name


# 不比较、不同步的列（期货表特有或不可覆盖）
_UNSYNCED_HEADERS = {"合同日期", "客户名称", "訂單編號", "鋼捲編號"}
# 日期列，比较时先归一化
_DATE_HEADERS = {"移撥日期", "入庫日期"}


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
            yehui_coil_info: 钢卷号 -> {order_no, warehouse, transfer_date, entry_date, data}
            schedule_data: 订单号 -> ScheduleRecord（用于补充客户信息）
        Returns:
            (preserved_rows, updated_coils, new_coils)
            updated_coils 中的元素包含 change_type 字段："transferred"（转单，紫色）或 "updated"（属性更新，橙色）
        """
        preserved: List[CoilRecord] = []
        updated: List[Dict[str, Any]] = []

        # 需要与烨辉同步的属性列（排除期货表特有列）
        sync_headers = [h for h in STANDARD_HEADERS if h not in _UNSYNCED_HEADERS]

        for row_info in original_rows:
            coil_no = row_info.coil_no
            if not coil_no or coil_no not in yehui_coil_info:
                # 烨辉表中不存在该钢卷 → 仍然保留（只增加不删减）
                preserved.append(row_info)
                continue

            yehui = yehui_coil_info[coil_no]
            new_order = yehui.get("order_no", "")
            old_order = row_info.order_no
            yehui_data = yehui.get("data", {})

            # 转单：烨辉订单号与原订单号不一致
            transferred = bool(new_order) and new_order != old_order

            changed = False
            if not transferred:
                for header in sync_headers:
                    old_val = self._get_row_value(row_info, header)
                    new_val = get_field_by_name(yehui_data, header)
                    if self._value_changed(header, old_val, new_val):
                        changed = True
                        break

            if transferred or changed:
                change_type = "transferred" if transferred else "updated"
                updated.append({
                    **row_info.to_dict(),
                    "order_no": new_order if transferred else old_order,
                    "change_type": change_type,
                    "yehui_data": yehui_data,
                    "new_order_no": new_order,
                    "old_order_no": old_order,
                })
            else:
                # 无任何变化 → 原样保留
                preserved.append(row_info)

        return preserved, updated, []

    def _get_row_value(self, row_info: CoilRecord, header: str) -> Any:
        """根据表头名从 row_data 中取值"""
        if header in FULL_HEADERS:
            idx = FULL_HEADERS.index(header)
            if idx < len(row_info.row_data):
                return row_info.row_data[idx]
        return None

    def _value_changed(self, header: str, old_val: Any, new_val: Any) -> bool:
        """比较字段是否变化，日期列先归一化"""
        if header in _DATE_HEADERS:
            return format_date(old_val) != format_date(new_val)
        old = str(old_val).strip() if old_val is not None else ""
        new = str(new_val).strip() if new_val is not None else ""
        return old != new

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
            updated_coils: 更新行（含 change_type: transferred/updated）
            new_coils: 新增行
            schedule_data: 用于补充转单行的客户信息
        Returns:
            customer_name -> [(type, data), ...]
        """
        customer_groups: Dict[str, List[Tuple[str, Any]]] = defaultdict(list)

        # 保留行
        for row_info in preserved_rows:
            customer = row_info.customer
            if customer:
                customer_groups[customer].append(("preserved", row_info))

        # 更新行（含转单与属性更新）
        for update_info in updated_coils:
            change_type = update_info.get("change_type", "updated")
            if change_type == "transferred":
                # 转单行：优先按新订单号从排程取客户归属
                order_no = update_info.get("new_order_no", "")
                customer = schedule_data.get(order_no, ScheduleRecord(order_no="", customer="")).customer
                if not customer:
                    customer = update_info.get("customer", "")
            else:
                customer = update_info.get("customer", "")
                if not customer:
                    order_no = update_info.get("order_no", "")
                    customer = schedule_data.get(order_no, ScheduleRecord(order_no="", customer="")).customer
            if customer:
                customer_groups[customer].append((change_type, update_info))

        # 新增行
        for new_info in new_coils:
            customer = new_info.get("customer", "")
            if customer:
                customer_groups[customer].append(("new", new_info))

        # 排序：按订单号排序，相同订单号内 preserved/updated/transferred 按原文件行号混合排列，
        # new 行追加在订单号末尾（保持添加顺序，依赖 Python sort 稳定性）
        for customer in customer_groups:
            customer_groups[customer].sort(
                key=lambda x: (
                    x[1].get("order_no", "") if isinstance(x[1], dict) else x[1].order_no,
                    0 if x[0] in ("preserved", "updated", "transferred") else 1,
                    x[1].get("row_idx", 999999) if isinstance(x[1], dict) else getattr(x[1], "row_idx", 999999),
                )
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
        """构建在客户 Sheet 中出现过的订单号集合（基于原文件）"""
        return {row.order_no for row in original_rows if row.order_no}

    @staticmethod
    def build_order_set_from_groups(customer_groups: Dict[str, List[Tuple[str, Any]]]) -> Set[str]:
        """构建从最终客户分组中实际写入的订单号集合"""
        orders: Set[str] = set()
        for items in customer_groups.values():
            for item_type, item_data in items:
                order_no = item_data.get("order_no", "") if isinstance(item_data, dict) else getattr(item_data, "order_no", "")
                if order_no:
                    orders.add(order_no)
        return orders
