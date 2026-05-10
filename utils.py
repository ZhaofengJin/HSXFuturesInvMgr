"""
通用工具模块
提供日期格式化、字段名匹配、文件查找等纯函数工具
"""

import os
from datetime import datetime

from config import FIELD_VARIANTS, SHEET_NAME_MAX_LEN, SHEET_NAME_INVALID_CHARS


def format_date(dt):
    """格式化日期为 YYYY/M/D 格式"""
    if dt is None:
        return ""
    if isinstance(dt, datetime):
        return f"{dt.year}/{dt.month}/{dt.day}"
    if isinstance(dt, str):
        try:
            dt_obj = datetime.strptime(dt.split()[0], "%Y-%m-%d")
            return f"{dt_obj.year}/{dt_obj.month}/{dt_obj.day}"
        except (ValueError, IndexError):
            return dt
    return str(dt)


def get_field_by_name(row_data, field_name):
    """
    从行数据中通过字段名获取值，支持字段名变体匹配
    Args:
        row_data: dict, 字段名 -> 值
        field_name: str, 标准字段名
    Returns:
        匹配到的值，未匹配返回 None
    """
    if field_name in row_data:
        return row_data[field_name]

    variants = FIELD_VARIANTS.get(field_name, [field_name])
    for variant in variants:
        if variant in row_data:
            return row_data[variant]

    return None


def safe_sheet_name(name):
    """
    将客户名称转换为安全的 Excel Sheet 名
    - 最长 31 字符
    - 替换非法字符为下划线
    """
    safe = name[:SHEET_NAME_MAX_LEN]
    for char in SHEET_NAME_INVALID_CHARS:
        safe = safe.replace(char, "_")
    return safe


def find_base_dir(search_dir, keyword):
    """
    在指定目录中查找包含关键字的子目录
    排除 ~$ 开头的临时目录
    Args:
        search_dir: str, 搜索目录
        keyword: str, 关键字
    Returns:
        str or None, 找到的目录路径
    """
    if not os.path.isdir(search_dir):
        return None

    for item in os.listdir(search_dir):
        if item.startswith("~$"):
            continue
        full_path = os.path.join(search_dir, item)
        if os.path.isdir(full_path) and keyword in item:
            return full_path

    return None


def is_temp_or_backup_file(filename):
    """判断是否为临时文件或备份文件"""
    basename = os.path.basename(filename)
    return (
        basename.startswith("~$")
        or "备份" in basename
        or ".bak" in basename.lower()
    )
