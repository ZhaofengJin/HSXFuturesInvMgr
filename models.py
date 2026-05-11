"""
数据模型模块
定义项目中使用的核心数据结构
"""

from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Any


@dataclass
class CoilRecord:
    """钢卷记录（来自原文件客户 Sheet）"""
    order_no: str
    coil_no: str
    warehouse: str = ""
    transfer_date: Any = None
    entry_date: Any = None
    customer: str = ""
    row_data: List[Any] = field(default_factory=list)
    fill: Any = None
    row_idx: int = 0
    modify_date: Any = None  # 修改日期（第24列），保留行需原样保留

    def to_dict(self) -> Dict[str, Any]:
        # 手动构建字典，避免 asdict 深拷贝 openpyxl 样式对象导致递归
        return {
            "order_no": self.order_no,
            "coil_no": self.coil_no,
            "warehouse": self.warehouse,
            "transfer_date": self.transfer_date,
            "entry_date": self.entry_date,
            "customer": self.customer,
            "row_data": self.row_data,
            "row_idx": self.row_idx,
            "modify_date": self.modify_date,
            # fill 是 openpyxl 对象，不序列化
        }


@dataclass
class ScheduleRecord:
    """期货排程记录"""
    order_no: str
    customer: str
    date: Any = None
    row: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "order_no": self.order_no,
            "customer": self.customer,
            "date": self.date,
            "row": self.row,
        }


@dataclass
class YehuiRecord:
    """烨辉库存表记录"""
    order_no: str
    coil_no: str
    data: Dict[str, Any] = field(default_factory=dict)

    @property
    def warehouse(self) -> str:
        from utils import get_field_by_name
        val = get_field_by_name(self.data, "倉別")
        return str(val).strip() if val else ""

    @property
    def transfer_date(self) -> Any:
        from utils import get_field_by_name
        return get_field_by_name(self.data, "移撥日期")

    @property
    def entry_date(self) -> Any:
        from utils import get_field_by_name
        return get_field_by_name(self.data, "入庫日期")


@dataclass
class ProcessingResult:
    """数据处理结果"""
    preserved: List[CoilRecord] = field(default_factory=list)
    updated: List[Dict[str, Any]] = field(default_factory=list)
    new: List[Dict[str, Any]] = field(default_factory=list)
    customer_groups: Dict[str, List[tuple]] = field(default_factory=dict)
    matched_count: int = 0
    unmatched_count: int = 0

    def summary(self) -> Dict[str, int]:
        return {
            "preserved": len(self.preserved),
            "updated": len(self.updated),
            "new": len(self.new),
            "matched": self.matched_count,
            "unmatched": self.unmatched_count,
        }
