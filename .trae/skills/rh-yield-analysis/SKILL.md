---
name: rh-yield-analysis
description: RH 测试数据良率分析 — 通道缺陷判定 32HD 动态统计限 Grade Map 与 Excel 报告导出. Use when the user asks to analyze RH yield data, run RH良率分析, check wafer/wafer yield, or generate RH Excel reports from test data folders. Do not use for unrelated manufacturing yield tools or non-RH test data.
---

# RH Yield Analysis

基于 Polars 的 RH 晶圆测试数据良率分析工具, 读取 Excel 测试文件, 按 DefectCode.xlsx 规格逐通道判定缺陷码, 汇总 Chip Grade 与各 Wafer 良率.

## 关键文件

| 文件 | 作用 |
|---|---|
| `rh_yield_gui.py` | 主程序 (Tkinter GUI + 分析逻辑 + Excel 导出), 约 1900 行 |
| `analyze_rh_yield.py` | 命令行 BJ% 区间统计脚本, 输出 CSV 与交互式 HTML |
| `merge_excel.py` | Bar / FPC 目录多文件合并脚本 |
| `yieldmap_gui.py` | 早期 Map 界面参考实现 |
| `DefectCode.xlsx` | 规格定义, 两个工作表: `32HD` 与 `S1S2` |

## 核心 API (`rh_yield_gui`)

```python
from rh_yield_gui import (
    load_defect_code,      # -> dict of spec groups, keys "32HD"/"S1S2"
    analyze,                # (source, spec_groups, only_r_chip, progress_cb) -> (result_dict, error_str)
    write_excel,            # (result_dict, save_path) 写出 Excel 报告
    CHANNEL_ORDER,          # ["S1"] + ["H1".. "H32"] + ["S2"]
    HD_CHANNELS,            # ["H1".. "H32"]
    SS_CHANNELS,            # ["S1", "S2"]
)
```

## 判定规则

- **通道顺序**: S1 + H1..H32 + S2
- **重复测试**: 同通道多文件/多行取 `Test time` 最新值
- **通道缺陷码**: 按 RejectPriority 升序合并, 首个不满足规格的码为最终结果
- **RejectPriority 为空**: 该参数不参与判定
- **32HD 动态限 (R / MR_Ratio)**: 按单个 Chip 的 32 个 HD 通道
  1. 3σ 剔除异常值 (ddof=1, 单次)
  2. R: mean ±20% 为单通道上下限
  3. MR_Ratio: mean ±35% 为单通道上下限, LSL ≥ 4
  4. 若 3σ 后样本标准差 = 0, 回退用全体样本均值
- **X+ / X- (Chip 级 R 均值越限)**:
  1. 取该 Chip 32HD R 的 3σ 剔除后均值 `fmean`
  2. 若 `fmean > R.USL` (静态上限, 默认 175) → 该 Chip 全部 HD 通道记为 `X+`
  3. 若 `fmean < R.LSL` (静态下限, 默认 95) → 该 Chip 全部 HD 通道记为 `X-`
  4. `X+`/`X-` 的 RejectPriority = `R.Priority - 1` (默认 0), 优先级高于 `R+`/`R-`, 会覆盖同 Chip 所有通道的其他缺陷码
  5. 被 `X+`/`X-` 标记的 Chip 整 Chip 数据 (含 S1/S2) 不参与 Density 密度图与 3σ 统计
- **Chip Grade**:
  - A = 32HD 全 OK 且 S1&S2 OK
  - B = S1&S2 OK 且 32HD OK 数 ∈ {30, 31}
  - 其余 NG, 报告根因 (X Chip 根因为 `X+`/`X-`)

## Excel 报告 (write_excel)

9 个工作表:

| Sheet | 内容 |
|---|---|
| Grade Summary | 每 Wafer Chip 数 / A/B/NG 计数 / Grade 良率 |
| Chip Detail | 每 Chip 各通道 Grade / dominant defect |
| Spec Used | 本次实际使用的规格 (含动态限计算结果) |
| Channel Yield | 按通道聚合的良率与主导缺陷码 |
| Chip DefectMap | 按 Chip 汇总的 Grade/主导缺陷热图 |
| Channel DefectMap | 按通道聚合的主导缺陷热图 (与上一等大 1:1) |
| Density | 32HD / S1S2 × R / MR_Ratio / ASYM / BJMP2% 共 8 张 KDE 密度图, LSL/USL 红色虚线 + 3σ 统计标注 |
| _density_data | 密度图隐藏数据表 |
| Date Trend | 按测试日期聚合的良率与缺陷率折线图 |

导出文件名默认 = `{目录名}_{YYYYMMDD_HHMMSS}.xlsx`.

## GUI 入口

```python
# 直接运行
python rh_yield_gui.py

# 或打包成单文件 exe:
pyinstaller --onefile --windowed --name RH_Yield_Analysis \
  --add-data "DefectCode.xlsx;." \
  --collect-all polars rh_yield_gui.py
```

## 给 Agent 的操作指引

1. 用户说"分析 RH 数据"时, 先确认数据源目录路径. 若路径是共享盘, 提醒首次扫描可能较慢 (程序用 `os.scandir` 遍历)
2. 若只有命令行, 用 `analyze(source, groups, only_r_chip=True, progress_cb=print)` 跑一次, 再 `write_excel` 导出
3. 导出 Excel 后可直接读各 Sheet 汇总良率与缺陷分布, 或打开 Density 表查看规格边界
4. 常见问题: `tuple.lstrip` 错通常是遇到 DefectCode.xlsx 中未列出的缺陷码 — 当前版本已在 `defect_color()` 归一化为 hex; 若再出现, 检查 openpyxl 样式前的数据类型
5. GUI 里 32HD 与 S1S2 的坐标轴范围已独立设定, 不要混用
