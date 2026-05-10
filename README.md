# HSXFuturesInvMgr

HSX 的期货库存管理工具（v5.0 TDD 重构版）

## 项目结构

```
HSXFuturesInvMgr/
├── main.py                  # 主处理脚本（v5.0 入口）
├── config.py                # 配置常量（颜色、表头、路径等）
├── utils.py                 # 通用工具（日期格式化、字段匹配等）
├── models.py                # 数据模型（CoilRecord, ScheduleRecord 等）
├── excel_handler.py         # Excel 读写封装
├── processor.py             # 核心业务逻辑（数据比对、分组等）
├── 启动器.py                 # 一键启动脚本
├── 使用说明.md               # 详细使用说明
├── README.md                # 本文件
├── tests/                   # 单元测试目录
│   ├── test_utils.py
│   ├── test_models.py
│   ├── test_excel_handler.py
│   └── test_processor.py
└── results/                 # 数据文件夹
    ├── 期货库存明细.xlsx
    ├── 期货库存明细_备份.xlsx
    └── 烨辉库存表.xlsx
```

## 核心功能

1. **备份原文件**：程序启动前自动备份 `期货库存明细.xlsx`
2. **读取期货排程**：提取订单号、客户名称、合同日期
3. **读取烨辉库存表**：建立钢卷号 → 仓别/移拨日期/入库日期 映射
4. **差异比对**：以期货排程订单为基准，识别新增/更新/保留
5. **颜色标注**：
   - 🟠 橙色：仓别已更新
   - 🟡 黄色：新增钢卷
   - 🟢 绿色：期货排程订单已匹配
   - 🔴 红色：期货排程订单未匹配
6. **只增加不删减**：原文件中已有的钢卷永远不会被删除

## 使用方法

### 命令行运行

```bash
# 使用 Python 直接运行
python main.py

# 或使用启动器
python 启动器.py
```

### 运行前准备

1. ✅ 关闭所有打开的 Excel 文件（尤其是 `期货库存明细.xlsx`）
2. ✅ 确认 `results/` 文件夹中有最新的 `期货库存明细.xlsx` 和 `烨辉库存表.xlsx`
3. ✅ 确认 `期货库存明细.xlsx` 中有 `期货排程` Sheet

## 运行测试

项目采用 **TDD（测试驱动开发）** 模式，所有核心逻辑均有单元测试覆盖。

```bash
# 运行全部测试
python -m pytest tests/ -v

# 或
python -m unittest discover tests/ -v
```

## 技术依赖

| 依赖 | 版本 |
|------|------|
| Python | 3.12+ |
| openpyxl | 3.1.5+ |
| pytest | 9.0+（可选，用于运行测试）|

## 版本历史

### v5.0 (2026-05-10) — TDD 重构版

- ✅ **代码重构**：将 818 行单体脚本拆分为 6 个独立模块
- ✅ **TDD 开发**：57 个单元测试覆盖所有核心逻辑
- ✅ **消除硬编码**：路径和配置提取到 `config.py`
- ✅ **跨平台支持**：支持 Windows 和 macOS/Linux
- ✅ **数据模型化**：使用 dataclass 定义核心数据结构
- ✅ **纯函数工具**：`utils.py` 提供无副作用的通用函数

### v4.5 及更早版本

详见 `使用说明.md`
