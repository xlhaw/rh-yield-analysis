# RH Yield Analysis — RH 测试数据良率分析工具

读取 RH 晶圆测试数据（Excel），按 DefectCode 规格逐通道判定缺陷码，汇总 Chip 等级（Grade）与各 Wafer 良率，提供图形界面查看与 Excel 报告导出。

## 功能

- **数据读取**：支持文件夹（自动扫描 .xlsx/.xls）或多选文件；网络共享盘可用；同一通道重复测试自动取最新 Test time
- **通道判定**：参数按 RejectPriority 升序合并，首个不满足规格的缺陷码作为通道最终结果；顺序为 S1 + H1~H32 + S2
- **32HD 动态统计限**：R / MR_Ratio 按单个 Chip 的 32 个 HD 通道 3σ 剔除后均值 ±20% / ±35%（下限 ≥4）动态计算
- **Grade 判定**：A（32HD 全 OK 且 S1/S2 OK）；B（S1/S2 OK 且 32HD OK 数为 30/31）；其余 NG 并给出根因
- **可视化**：by-Chip Grade Map、by-Channel Defect Map（等大 1:1 并排）、两个饼图
- **Excel 报告**：Grade Summary / Chip Detail / Spec Used / Channel Yield / 两个 DefectMap / Density（8 张 KDE 密度图，含 LSL/USL 红色虚线与 3σ 统计标注）/ Date Trend

## 环境要求

- Windows 10/11（64 位）
- 发布版为单文件 exe（`RH_Yield_Analysis.exe`），无需 Python 环境
- 从源码运行需要：Python 3.9+、polars、matplotlib、numpy、openpyxl、pandas、xlrd

## 使用

1. 将 `DefectCode.xlsx` 与程序放同一目录（或在界面中指定路径）
2. 双击运行 `RH_Yield_Analysis.exe`，选择数据源（文件夹或文件）
3. 点击「开始分析」，完成后在右侧查看 Map 与饼图，可按 Wafer 切换
4. 点击「导出 Excel」生成报告，默认文件名含数据源目录名与生成时间戳

详细操作与判定规则见 [RH_Yield_Analysis_操作手册.docx](RH_Yield_Analysis_操作手册.docx)。

## 文件说明

| 文件 | 说明 |
|---|---|
| `rh_yield_gui.py` | 主程序（Tkinter GUI + 分析 + 导出） |
| `analyze_rh_yield.py` | 命令行统计脚本（BJ% 区间统计，输出 CSV 与交互式 HTML） |
| `merge_excel.py` | 数据合并脚本（Bar / FPC 目录多文件合并） |
| `yieldmap_gui.py` | 早期 Map 界面参考实现 |
| `DefectCode.xlsx` | 规格定义（32HD / S1S2 两个工作表） |

## 打包

```bash
pyinstaller --onefile --windowed --name RH_Yield_Analysis \
  --add-data "DefectCode.xlsx;." \
  --collect-all polars rh_yield_gui.py
```

## 注意

- 仓库中的脚本默认路径均为占位符（`\\server\share\Raw data`），使用时请改为实际数据源路径
- `rh_yield_config.json` 记录本机数据源路径，已加入 `.gitignore`，不会提交
