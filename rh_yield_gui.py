# -*- coding: utf-8 -*-
"""
RH 测试数据良率分析 GUI (基于 Polars)

功能:
  1. 读取 RH 测试数据 xlsx/xls 的 'R' sheet (每文件一颗 Chip, 34 通道: S1, H1~H32, S2),
     列映射参照 merge_excel.py 的统一列方式。
  2. 参照 DefectCode.xlsx 的 '32HD' / 'S1S2' sheet 读取各项参数的 Spec (LSL/USL) 和
     RejectPriority。RejectPriority 为空 ('/' 或空白) 的参数不参与良率判断。
  3. 参照 yieldmap_gui.py 的 defect_map / merge_defect_priority 方法:
     按各项参数的 Reject Priority 依次判定 (低于 LSL 记 LSL_Code, 高于 USL 记 USL_Code),
     并按优先级合并出每个通道的 Map (首个非 OK 的缺陷码)。
     - 32HD sheet 的 BJMP% / BJMP2% 在 DefectCode 中未给定 Spec,
       可在界面手动填写 (留空则不判定)。
  4. Chip Grade 判定:
     - A: 32HD 全部 OK 且 S1 & S2 均 OK
     - B: S1 & S2 均 OK 且 32HD OK 数为 31 或 30 (共 32)
     - NG: 其余
  5. 输出: 每片晶圆的 Grade Map (A/B/NG), 良率汇总, NG 根因分布, Excel 导出。
"""
import re
import sys
import json
from datetime import datetime
from pathlib import Path

import polars as pl
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.patches as mpatches
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.text import (Paragraph, ParagraphProperties,
                                   CharacterProperties, RegularTextRun)
import numpy as np

# PyInstaller 打包后: BASE_DIR = exe 同目录 (可写, 存放 config 等);
# RES_DIR = 打包资源目录 (_MEIPASS: onefile 解压临时目录 / onedir 的 _internal)。
# 未打包时两者均为脚本所在目录。
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
    RES_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
else:
    BASE_DIR = Path(__file__).parent
    RES_DIR = BASE_DIR


def _find_resource(name):
    """优先使用 exe 同目录的外部文件 (可编辑), 否则回退到打包内置资源。"""
    ext = BASE_DIR / name
    if ext.exists():
        return ext
    return RES_DIR / name


CONFIG_FILE = BASE_DIR / "rh_yield_config.json"
DEFECT_CODE_FILE = _find_resource("DefectCode.xlsx")

# 通道定义: 32HD = H1~H32 (数据中为 'H1'..'H32'), S1S2 = 'S1','S2'
HD_CHANNELS = [f"H{i}" for i in range(1, 33)]
SS_CHANNELS = ["S1", "S2"]
CHANNEL_ORDER = ["S1"] + HD_CHANNELS + ["S2"]

IDENT_COLS = ["Wafer ID", "Part ID", "Hd", "Test time"]

GRADE_COLORS = {"A": "#4CAF50", "B": "#FF9800", "NG": "#F44336"}
GRADE_ORDER = ["A", "B", "NG"]

# 缺陷码 -> 颜色 (用于 by-Channel DefectMap)
DEFECT_COLORS = {
    "OK": "#4CAF50",
    "R-": "#E57373", "R+": "#EF5350",
    "X-": "#7B1FA2", "X+": "#6A1B9A",
    "MR-": "#FF8A65", "MR+": "#FF7043",
    "A-": "#FFB74D", "A+": "#FFA726",
    "PIN": "#9575CD", "SPIN": "#7E57C2",
    "BJ-": "#4FC3F7", "BJ+": "#29B6F6",
    "BJ2-": "#81D4FA", "BJ2+": "#4FC3F7",
    "BDV-": "#AED581", "BDV+": "#9CCC65",
    "SR-": "#BA68C8", "SR+": "#AB47BC",
    "SMR-": "#F06292", "SMR+": "#EC407A",
    "SBJ+": "#4DD0E1",
    "HY-": "#FFAB91", "HY+": "#FF8A65",
    "SHY-": "#CE93D8",
}
_DEFECT_CMAP = plt.get_cmap("tab20")


# ── 配置持久化 ─────────────────────────────────────────────

def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ── DefectCode 读取 ────────────────────────────────────────

def _clean_cell(v):
    """'/' / 空白 / None / NaN -> None; 其余原样返回。"""
    if v is None:
        return None
    if isinstance(v, float) and v != v:  # NaN
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "/"):
            return None
        return s
    return v


def load_defect_code(path):
    """读取 DefectCode.xlsx 的 '32HD' / 'S1S2' sheet。

    返回 {group: [spec, ...]}，
    spec = {ColumnName, ShortName, Priority(int|None), LSL, USL,
            LSL_Code, USL_Code, Statistic}
    Priority / LSL / USL 解析失败或为空记 None。
    """
    df_spec = pl.read_excel(str(path), sheet_name="32HD", engine="openpyxl")
    cols = {c.strip(): c for c in df_spec.columns}
    need = ["ColumnName", "ShortName", "RejectPriority", "LSL", "USL",
            "LSL_Code", "USL_Code", "Statistic"]
    for n in need:
        if n not in cols:
            raise ValueError(f"DefectCode 缺少列: {n}")

    def num(v):
        v = _clean_cell(v)
        if v is None:
            return None
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    def prio(v):
        v = _clean_cell(v)
        if v is None:
            return None
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None

    def read_sheet(sheet_name):
        df = pl.read_excel(str(path), sheet_name=sheet_name, engine="openpyxl")
        header = list(df.columns)
        idx = {n: header.index(cols[n]) for n in need}
        rows = []
        for row in df.rows():
            cn = _clean_cell(row[idx["ColumnName"]])
            if cn is None:
                continue
            rows.append({
                "ColumnName": cn,
                "ShortName": _clean_cell(row[idx["ShortName"]]) or cn,
                "Priority": prio(row[idx["RejectPriority"]]),
                "LSL": num(row[idx["LSL"]]),
                "USL": num(row[idx["USL"]]),
                "LSL_Code": _clean_cell(row[idx["LSL_Code"]]),
                "USL_Code": _clean_cell(row[idx["USL_Code"]]),
                "Statistic": _clean_cell(row[idx["Statistic"]]),
            })
        return rows

    return {"32HD": read_sheet("32HD"), "S1S2": read_sheet("S1S2")}


def active_params(spec_rows):
    """RejectPriority 非空的参数, 按 Priority 升序。"""
    act = [s for s in spec_rows if s["Priority"] is not None]
    return sorted(act, key=lambda s: s["Priority"])


# ── 数据读取 (参照 merge_excel.py) ─────────────────────────

_FILE_WAFER_CHIP = re.compile(r"^([^_]+)_([^_]+)_")


def collect_data_files(source):
    """source 为文件夹时递归收集 xlsx/xls (跳过 ~$ 临时文件), 否则原样返回列表。

    使用 os.scandir 而非 pathlib.rglob: DirEntry.is_file() 利用目录项 d_type
    缓存, 避免对每个文件单独 stat (网络盘上 stat 极慢)。
    """
    import os
    p = str(source)
    if os.path.isfile(p):
        return [p]
    if not os.path.isdir(p):
        return []
    exts = (".xlsx", ".xls")
    out = []
    stack = [p]
    while stack:
        cur = stack.pop()
        try:
            with os.scandir(cur) as it:
                for e in it:
                    if e.name.startswith("~$"):
                        continue
                    try:
                        if e.is_dir(follow_symlinks=False):
                            stack.append(e.path)
                        elif e.is_file(follow_symlinks=False):
                            if e.name.lower().endswith(exts):
                                out.append(e.path)
                    except OSError:
                        continue
        except OSError:
            continue
    out.sort()
    return out


# ── 判定与 Map 合并 (参照 yieldmap_gui.py) ─────────────────

def build_ng_expr(col, spec, lsl_expr=None, usl_expr=None):
    """单参数判定表达式: 返回缺陷码或 'OK'。

    值 < LSL -> LSL_Code; 值 > USL -> USL_Code (对应侧 Code 为空则该侧不判定);
    值缺失 -> ShortName (数据缺失按 NG 处理)。

    lsl_expr / usl_expr: 可选, 用 polars 列表达式替代 spec 中的标量 LSL/USL
                          (用于 32HD 按 Chip 动态限)。
    """
    v = pl.col(col)
    conds = []
    lsl_used = lsl_expr if lsl_expr is not None else (
        pl.lit(spec["LSL"]) if spec["LSL"] is not None else None)
    usl_used = usl_expr if usl_expr is not None else (
        pl.lit(spec["USL"]) if spec["USL"] is not None else None)
    if lsl_used is not None and spec["LSL_Code"]:
        conds.append((v < lsl_used, spec["LSL_Code"]))
    if usl_used is not None and spec["USL_Code"]:
        conds.append((v > usl_used, spec["USL_Code"]))
    conds.append((v.is_null(), f"{spec['ShortName']}(NA)"))
    expr = pl.lit("OK", dtype=pl.String)
    for cond, code in reversed(conds):
        expr = pl.when(cond).then(pl.lit(code)).otherwise(expr)
    return expr


def compute_dynamic_limits_32hd(data, specs, x_r_limits=None):
    """按 Chip 计算 32HD 'R' 与 'MR_Ratio' 的动态上下限。

    每个 Chip 用自身 32 个 HD 通道数据: 先剔除 3*Sigma 外值, 求均值,
    R → 均值±20%, MR_Ratio → 均值±35% (LSL≥4)。
    全部相同/为空时 LSL/USL 留空 (回退 DefectCode 标量)。

    对 R 额外计算 X+/X-: 若 3σ 剔除后均值低于静态 R LSL 记 'X-',
    高于静态 R USL 记 'X+', 否则 None。X+/- 优先级高于 R+/-。
    x_r_limits: {'LSL': float|None, 'USL': float|None} 覆盖 X+/- 判定
                用的静态 R 上下限; 为 None 时用 DefectCode 中的值。

    返回 DataFrame: WaferID, ChipID, {col}__LSL, {col}__USL,
                    {col}__mean, {col}__keep_n, {col}__n, {col}__X_code (仅 R)
    """
    hd = data.filter(pl.col("Hd").is_in(HD_CHANNELS))
    target = {s["ColumnName"]: s for s in specs if s["ShortName"] in ("R", "MR_Ratio")}
    if not target:
        return pl.DataFrame(schema={"WaferID": pl.String, "ChipID": pl.String})

    # 逐 Chip 计算 (用 partition_by 避免 map_groups 的分组表达式限制)
    rows = []
    for g in hd.partition_by(["WaferID", "ChipID"], as_dict=False):
        row = {"WaferID": g["WaferID"][0], "ChipID": g["ChipID"][0]}
        for col, spec in target.items():
            lsl_c, usl_c = f"{col}__LSL", f"{col}__USL"
            mean_c, keep_c, n_c = (f"{col}__mean", f"{col}__keep_n", f"{col}__n")
            if col not in g.columns:
                row.update({lsl_c: None, usl_c: None, mean_c: None,
                            keep_c: 0, n_c: 0})
                continue
            vals = g[col].drop_nulls()
            n = len(vals)
            if n == 0:
                row.update({lsl_c: None, usl_c: None, mean_c: None,
                            keep_c: 0, n_c: 0})
                continue
            mean0 = vals.mean()
            std0 = vals.std()
            if std0 is None or std0 == 0:
                fmean, keep_n = mean0, n
            else:
                lo, hi = mean0 - 3 * std0, mean0 + 3 * std0
                filtered = vals.filter((vals >= lo) & (vals <= hi))
                keep_n = len(filtered)
                fmean = filtered.mean() if keep_n > 0 else mean0
            if spec["ShortName"] == "R":
                lsl, usl = fmean * 0.80, fmean * 1.20
                # X+/-: 3σ 剔除后均值超过静态 R 规格上下限 (GUI 可覆盖)
                if x_r_limits:
                    s_lsl = x_r_limits.get("LSL") or spec["LSL"]
                    s_usl = x_r_limits.get("USL") or spec["USL"]
                else:
                    s_lsl, s_usl = spec["LSL"], spec["USL"]
                x_code = None
                if s_lsl is not None and fmean < s_lsl:
                    x_code = "X-"
                elif s_usl is not None and fmean > s_usl:
                    x_code = "X+"
                row[f"{col}__X_code"] = x_code
            else:
                lsl, usl = max(fmean * 0.65, 4.0), fmean * 1.35
            row.update({lsl_c: lsl, usl_c: usl, mean_c: fmean,
                        keep_c: int(keep_n), n_c: int(n)})
        rows.append(row)
    # X_code 列多数行为 None、少数为 'X+'/'X-'，polars 按前 N 行推断会误判为
    # Null 类型导致后续 append 失败。infer_schema_length=None 强制扫全表。
    return pl.DataFrame(rows, infer_schema_length=None) if rows else pl.DataFrame(
        schema={"WaferID": pl.String, "ChipID": pl.String})


def judge_group(df, specs, extra_limits=None, limit_override_cols=None):
    """对一组通道数据按 Reject Priority 判定并合并 Map。

    df: 每行一个通道测量 (WaferID, ChipID, Hd + 参数列)
    specs: DefectCode 某组 (32HD/S1S2) 的参数行
    extra_limits: {ColumnName: (LSL|None, USL|None)} 手动补充的 Spec (仅当 DefectCode 无 Spec 时生效)
    limit_override_cols: {ColumnName: (lsl_col_name, usl_col_name)}
                          按行使用 df 中的限列替代 DefectCode 标量 LSL/USL (32HD 按 Chip 动态限)
    返回 (df 增加 NG 列 + 'CODE' 最终 Map + '_PRIO' 缺陷优先级, code->priority dict)
    """
    extra_limits = extra_limits or {}
    limit_override_cols = limit_override_cols or {}
    act = active_params(specs)
    code_prio = {}
    df = df.with_columns(pl.lit("OK", dtype=pl.String).alias("CODE"))
    df = df.with_columns(pl.lit(None, dtype=pl.Int64).alias("_PRIO"))

    for s in act:
        col = s["ColumnName"]
        if col not in df.columns:
            # 整列缺失 -> 无法判定, 跳过 (结果面板会提示)
            continue
        spec = dict(s)
        lsl_expr = usl_expr = None
        # 1. 按 Chip 动态限: 用 df 中的限列覆盖标量
        if col in limit_override_cols:
            lsl_col, usl_col = limit_override_cols[col]
            lsl_expr = pl.col(lsl_col)
            usl_expr = pl.col(usl_col)
        # 2. 手动补充 Spec: 仅当 DefectCode 未给定 LSL/USL 时生效
        elif spec["LSL"] is None and spec["USL"] is None:
            extra = extra_limits.get(col) or extra_limits.get(s["ShortName"])
            if extra:
                el_lsl, el_usl = extra
                if el_lsl is not None:
                    spec["LSL"] = el_lsl
                    if not spec["LSL_Code"]:
                        spec["LSL_Code"] = f"{s['ShortName']}-"
                if el_usl is not None:
                    spec["USL"] = el_usl
                    if not spec["USL_Code"]:
                        spec["USL_Code"] = f"{s['ShortName']}+"
        # 无 Spec 且无动态限/手动限 -> 不判定
        if (lsl_expr is None and usl_expr is None
                and spec["LSL"] is None and spec["USL"] is None):
            continue
        ng_col = f"_NG_{s['Priority']}_{s['ShortName']}"
        df = df.with_columns(
            build_ng_expr(col, spec, lsl_expr=lsl_expr, usl_expr=usl_expr).alias(ng_col))
        # Map 合并: 按 Priority 升序, 首个非 OK 生效
        df = df.with_columns(
            pl.when(pl.col("CODE") == "OK")
            .then(pl.col(ng_col))
            .otherwise(pl.col("CODE"))
            .alias("CODE")
        )
        df = df.with_columns(
            pl.when(pl.col("_PRIO").is_null() & (pl.col(ng_col) != "OK"))
            .then(pl.lit(s["Priority"]))
            .otherwise(pl.col("_PRIO"))
            .alias("_PRIO")
        )
        code_prio[s["Priority"]] = s["ShortName"]
    return df, code_prio


def parse_chip_rowcol(chip_id):
    """'050R04' -> (50, 4); 解析失败 -> (None, None)。"""
    m = re.match(r"(\d+)[RW](\d+)", str(chip_id), re.IGNORECASE)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


# ── 主分析流程 ─────────────────────────────────────────────

def parse_lot_wafer(source, files):
    """从数据源目录名解析 LotNo / WaferID 及导出用文件名。

    目录名形如 '268082-J4_TR_MD26A0407' / '259052 MD24G0212' /
    '25B112_TW_MD24I0180'。分隔符 (空格/-/_) 统一归一为 '_'；
    LotNo 取第一段, WaferID 取以 'MD' 开头的末段 (无匹配时取末段)。
    返回 (stem, lot_no, wafer_id)；无法解析时返回 (None, None, None)。
    """
    import os
    d = str(source or "")
    if os.path.isfile(d):
        # 源为多个文件 (以 ';' 分隔) 时取第一个文件的所在目录
        d = d.split(";")[0]
    if os.path.isfile(d):
        d = os.path.dirname(d)
    if not os.path.isdir(d):
        if not files:
            return None, None, None
        d = os.path.dirname(str(files[0]))
    name = os.path.basename(os.path.normpath(d))
    tokens = [t for t in re.split(r"[ _\-]+", name) if t]
    if not tokens:
        return None, None, None
    lot_no = tokens[0]
    wafer_id = next((t for t in reversed(tokens) if t.upper().startswith("MD")),
                    tokens[-1])
    return "_".join(tokens), lot_no, wafer_id


def analyze(source, spec_groups, bj_limits=None, only_r_chip=True,
            x_r_limits=None, progress_cb=None):
    """分析数据源, 返回 result dict。

    bj_limits: {'32HD': {'BJMP%': (None, UL), 'BJMP2%': (None, UL)}}
               DefectCode 未给定 Spec 的 32HD BJ 参数手动上限。
    x_r_limits: {'LSL': float|None, 'USL': float|None}
                覆盖 X+/- 判定用的静态 R 上下限; 为 None 时用 DefectCode 值。
    """
    bj_limits = bj_limits or {}
    files = collect_data_files(source)
    if not files:
        return None, "未找到数据文件 (xlsx/xls)"

    all_params = ({s["ColumnName"] for g in ("32HD", "S1S2") for s in spec_groups[g]}
                  - {"Wafer ID", "Part ID", "Hd"})
    # 预期读取的列: IDENT_COLS + 所有参数列 (用于 calamine 只读需要的列)
    _want_cols = list(dict.fromkeys(list(IDENT_COLS) + sorted(all_params)))

    def _read_one(f):
        ext = f.lower()
        engine = "xlrd" if ext.endswith(".xls") else "calamine"
        try:
            raw = pl.read_excel(f, sheet_name="R", engine=engine, columns=_want_cols)
        except Exception:
            # columns 参数失败 (列名不匹配等) 时回退: 全读再筛
            try:
                raw = pl.read_excel(f, sheet_name="R", engine=engine)
            except Exception:
                if engine == "calamine":
                    raw = pl.read_excel(f, sheet_name="R", engine="openpyxl")
                else:
                    raise
        raw.columns = [str(c).strip() for c in raw.columns]
        keep = [c for c in IDENT_COLS if c in raw.columns]
        param_present = [c for c in sorted(all_params) if c in raw.columns]
        df = raw.select(keep + param_present)
        df = df.rename({k: v for k, v in
                        {"Wafer ID": "WaferID", "Part ID": "ChipID",
                         "Hd": "Hd", "Test time": "TestTime"}.items()
                        if k in df.columns})
        for c in ("WaferID", "ChipID", "Hd", "TestTime"):
            if c not in df.columns:
                df = df.with_columns(pl.lit(None, dtype=pl.String).alias(c))
        m = _FILE_WAFER_CHIP.match(Path(f).name)
        fw, fc = (m.group(1), m.group(2)) if m else (None, None)
        # WaferID / ChipID 始终以文件名为准 —— xlsx 内部 OCR/手录常出错
        # (如 '146R06' 被误读为 '1A6R06', '5L0506' 被误录为数字 '509506')
        df = df.with_columns([
            pl.lit(fw).alias("WaferID"),
            pl.lit(fc).alias("ChipID"),
        ])
        df = df.with_columns(
            pl.col("WaferID").cast(pl.String).str.strip_chars().alias("WaferID"),
            pl.col("ChipID").cast(pl.String).str.strip_chars().alias("ChipID"),
            pl.col("Hd").cast(pl.String).str.strip_chars().alias("Hd"),
            pl.col("TestTime").cast(pl.String).alias("TestTime"),
        )
        df = df.with_columns(pl.lit(Path(f).name).alias("_FILE"))
        return df

    frames = []
    errors = []
    nf = len(files)
    # 多线程读取 (calamine 解析时释放 GIL, 线程可并行)
    from concurrent.futures import ThreadPoolExecutor
    import threading
    _lock = threading.Lock()
    done = [0]

    def _task(idx_f):
        i, f = idx_f
        try:
            d = _read_one(f)
            with _lock:
                frames.append(d)
                done[0] += 1
                if progress_cb and done[0] % 20 == 0:
                    progress_cb(done[0], nf)
        except Exception as e:
            with _lock:
                errors.append(f"{Path(f).name}: {e}")
                done[0] += 1

    with ThreadPoolExecutor(max_workers=8) as ex:
        list(ex.map(_task, enumerate(files, 1)))

    if not frames:
        return None, "没有可读取的 'R' sheet 数据\n" + "\n".join(errors[:10])

    data = pl.concat(frames, how="diagonal_relaxed")

    # 仅统计 ChipID 含 'R' 的 Chip (参照 analyze_rh_yield.py)
    if only_r_chip:
        data = data.filter(pl.col("ChipID").str.contains("R"))

    if data.height == 0:
        return None, "过滤后无数据 (ChipID 含 'R')"

    # 同一 (WaferID, ChipID, Hd) 多条记录时保留最新 TestTime (复测取最新)
    data = data.with_columns(
        pl.col("TestTime").str.to_datetime(strict=False, exact=False).alias("_TS")
    )
    data = data.sort("_TS", nulls_last=False).unique(
        subset=["WaferID", "ChipID", "Hd"], keep="last", maintain_order=True)

    # 参数列统一为 Float64
    pcols = [c for c in data.columns if c in all_params]
    data = data.cast({c: pl.Float64 for c in pcols})

    # 分组判定
    hd_data = data.filter(pl.col("Hd").is_in(HD_CHANNELS))
    ss_data = data.filter(pl.col("Hd").is_in(SS_CHANNELS))

    # 32HD 'R' / 'MR_Ratio' 动态限 (按 Chip: 每 Chip 自身 32 个 HD 通道统计)
    dynamic_limits = compute_dynamic_limits_32hd(
        data, spec_groups["32HD"], x_r_limits=x_r_limits)

    # 将按 Chip 的限列 join 到 hd_data, 单次判定 (向量化)
    limit_cols = {}
    if dynamic_limits.height > 0:
        hd_data = hd_data.join(dynamic_limits, on=["WaferID", "ChipID"], how="left")
        target = {s["ColumnName"]: s for s in spec_groups["32HD"]
                  if s["ShortName"] in ("R", "MR_Ratio")}
        for col in target:
            lsl_col = f"{col}__LSL"
            usl_col = f"{col}__USL"
            if lsl_col in dynamic_limits.columns:
                limit_cols[col] = (lsl_col, usl_col)

    extra_hd = bj_limits.get("32HD", {})
    hd_judged, hd_code_prio = judge_group(
        hd_data, spec_groups["32HD"],
        extra_limits=extra_hd, limit_override_cols=limit_cols)
    ss_judged, ss_code_prio = judge_group(ss_data, spec_groups["S1S2"])

    # ── X+/- chip 级覆盖 (优先级在 R+/- 之前) ──
    # 当 Chip 的 32HD R 3σ 剔除后均值超过静态 R 规格上下限时,
    # 该 Chip 全部 HD 通道记为 X+/X-, 优先级 = R.Priority - 1。
    r_spec_x = next((s for s in spec_groups["32HD"] if s["ShortName"] == "R"), None)
    x_col = f"{r_spec_x['ColumnName']}__X_code" if r_spec_x else None
    if (dynamic_limits.height > 0 and x_col
            and x_col in dynamic_limits.columns):
        x_df = dynamic_limits.filter(pl.col(x_col).is_not_null())
        if x_df.height > 0:
            x_prio = (r_spec_x["Priority"] - 1) if r_spec_x["Priority"] is not None else 0
            x_map = x_df.select(["WaferID", "ChipID", x_col])
            hd_judged = hd_judged.join(x_map, on=["WaferID", "ChipID"], how="left")
            cond = (pl.col(x_col).is_not_null()
                    & (pl.col("_PRIO").is_null() | (pl.col("_PRIO") > x_prio)))
            hd_judged = hd_judged.with_columns(
                pl.when(cond).then(pl.col(x_col))
                .otherwise(pl.col("CODE")).alias("CODE"),
                pl.when(cond).then(pl.lit(x_prio))
                .otherwise(pl.col("_PRIO")).alias("_PRIO"),
            ).drop(x_col)
            hd_code_prio = {**hd_code_prio, x_prio: "X"}

    code_prio = {**hd_code_prio, **ss_code_prio}

    judged = pl.concat([hd_judged, ss_judged], how="diagonal_relaxed")

    # ── Chip 级 Grade 判定 ──
    # 通道 Map 透视: 每 Chip 一行, 列 = S1, H1..H32, S2 的最终缺陷码
    chan = judged.select(["WaferID", "ChipID", "Hd", "CODE"])
    pivot = chan.pivot(values="CODE", index=["WaferID", "ChipID"], on="Hd",
                       aggregate_function="first")

    def ch(c):
        if c in pivot.columns:
            return (pl.col(c).fill_null("~") == "OK")
        return pl.lit(False)  # 通道无数据 -> 不 OK

    hd_ok_exprs = [ch(c) for c in HD_CHANNELS]
    pivot = pivot.with_columns(
        pl.sum_horizontal(hd_ok_exprs).cast(pl.Int64).alias("HD_OK"))
    hd_n_exprs = [(pl.col(c).is_not_null() if c in pivot.columns else pl.lit(False))
                  for c in HD_CHANNELS]
    pivot = pivot.with_columns(
        pl.sum_horizontal(hd_n_exprs).cast(pl.Int64).alias("HD_N"))
    s1_ok = ch("S1")
    s2_ok = ch("S2")
    ss_ok = s1_ok & s2_ok

    pivot = pivot.with_columns(
        pl.when(ss_ok & (pl.col("HD_OK") == 32)).then(pl.lit("A"))
        .when(ss_ok & pl.col("HD_OK").is_in([30, 31])).then(pl.lit("B"))
        .otherwise(pl.lit("NG")).alias("GRADE")
    )

    # NG 根因: 所有失败通道中优先级最小 (数字最小) 的缺陷码
    fail = judged.filter(pl.col("CODE") != "OK").select(
        ["WaferID", "ChipID", "Hd", "CODE", "_PRIO"])
    root = (fail.sort("_PRIO", nulls_last=True)
            .group_by(["WaferID", "ChipID"]).first()
            .rename({"CODE": "ROOT_CODE", "_PRIO": "ROOT_PRIO", "Hd": "ROOT_CH"}))

    chips = pivot.join(root, on=["WaferID", "ChipID"], how="left")
    chips = chips.with_columns(
        pl.when(pl.col("GRADE") == "NG")
        .then(
            pl.when(pl.col("ROOT_CODE").is_not_null())
            .then(pl.col("ROOT_CODE"))
            .otherwise(pl.lit("MISSING"))
        )
        .otherwise(pl.lit(None, dtype=pl.String))
        .alias("ROOT_CODE")
    )

    # ── Wafer 汇总 ──
    wafer_sum = (chips.group_by("WaferID").agg(
        pl.len().alias("Total"),
        (pl.col("GRADE") == "A").sum().alias("A"),
        (pl.col("GRADE") == "B").sum().alias("B"),
        (pl.col("GRADE") == "NG").sum().alias("NG"),
    ).sort("WaferID"))
    wafer_sum = wafer_sum.with_columns([
        (pl.col("A") / pl.col("Total") * 100).round(2).alias("A%"),
        (pl.col("B") / pl.col("Total") * 100).round(2).alias("B%"),
        ((pl.col("A") + pl.col("B")) / pl.col("Total") * 100).round(2).alias("Pass%"),
    ])

    # 根因分布 (按 Wafer + 根因码)
    cause = (chips.filter(pl.col("GRADE") == "NG")
             .group_by(["WaferID", "ROOT_CODE"]).agg(pl.len().alias("CNT"))
             .sort("WaferID", "ROOT_CODE"))

    # 通道级明细 (长表): WaferID, ChipID, Hd, CODE + 各参数 NG 列
    ng_cols = [c for c in judged.columns if c.startswith("_NG_")]
    detail_params = judged.select(
        ["WaferID", "ChipID", "Hd", "CODE"] +
        [c for c in pcols if c in judged.columns])

    return {
        "files": files,
        "errors": errors,
        "judged": judged,
        "chips": chips,
        "wafer_summary": wafer_sum,
        "cause": cause,
        "detail_params": detail_params,
        "code_prio": code_prio,
        "spec_groups": spec_groups,
        "bj_limits": bj_limits,
        "x_r_limits": x_r_limits,
        "dynamic_limits": dynamic_limits,
        "source": str(source),
        "analyzed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_stem": parse_lot_wafer(source, files)[0],
    }, None


# ── Wafer Map 绘制 (参照 yieldmap_gui.py plot_wafer_map) ───

def generate_chip_grid(chips_df, wafer_id):
    sub = chips_df.filter(pl.col("WaferID") == wafer_id)
    rows, cols, grades, chipids = [], [], [], []
    for cid, g in zip(sub["ChipID"], sub["GRADE"]):
        r, c = parse_chip_rowcol(cid)
        if r is None:
            r, c = 0, len(rows) + 1
        rows.append(r)
        cols.append(c)
        grades.append(g)
        chipids.append(cid)
    if not rows:
        return None
    return {
        "grid": dict(zip(zip(rows, cols), grades)),
        "chipids": dict(zip(zip(rows, cols), chipids)),
        "r_min": min(rows), "r_max": max(rows),
        "c_min": min(cols), "c_max": max(cols),
    }


def plot_grade_map(ax, grid_data, title=""):
    if grid_data is None:
        ax.text(0.5, 0.5, "No Data", ha="center", va="center", fontsize=14)
        ax.set_title(title)
        return
    grid = grid_data["grid"]
    n_rows = grid_data["r_max"] - grid_data["r_min"] + 1
    n_cols = grid_data["c_max"] - grid_data["c_min"] + 1
    for r in range(grid_data["r_min"], grid_data["r_max"] + 1):
        for c in range(grid_data["c_min"], grid_data["c_max"] + 1):
            val = grid.get((r, c))
            color = GRADE_COLORS.get(val, "#E0E0E0")
            i, j = r - grid_data["r_min"], c - grid_data["c_min"]
            ax.add_patch(plt.Rectangle((j, n_rows - 1 - i), 1, 1,
                                       facecolor=color, edgecolor="white",
                                       linewidth=0.3))
    ax.set_xlim(0, n_cols)
    ax.set_ylim(0, n_rows)
    ax.set_aspect("auto")
    ax.set_box_aspect(1)
    ax.set_xlabel("ChipCol")
    ax.set_ylabel("ChipRow")
    ax.set_title(title, fontsize=10)
    ax.tick_params(labelsize=6)


def add_grade_legend(ax):
    handles = [mpatches.Patch(color=GRADE_COLORS[k], label=k) for k in GRADE_ORDER]
    handles.append(mpatches.Patch(color="#E0E0E0", label="Untested"))
    ax.legend(handles=handles, loc="upper right", fontsize=7, framealpha=0.8)


def plot_grade_pie(ax, chips_df, wafer_filter="ALL", title="By-Chip Grade"):
    """by-Chip DefectMap 下方的 Grade 分布饼图 (A/B/NG)。"""
    df = chips_df
    if wafer_filter and wafer_filter != "ALL":
        df = df.filter(pl.col("WaferID") == wafer_filter)
    counts = {g: 0 for g in GRADE_ORDER}
    for g in df["GRADE"].to_list():
        if g in counts:
            counts[g] += 1
    labels, sizes, colors = [], [], []
    for g in GRADE_ORDER:
        if counts[g] > 0:
            labels.append(f"{g}")
            sizes.append(counts[g])
            colors.append(GRADE_COLORS[g])
    if not sizes:
        ax.text(0.5, 0.5, "No Data", ha="center", va="center", fontsize=11)
        ax.set_title(title, fontsize=10)
        ax.axis("off")
        return
    total = sum(sizes)
    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors,
        autopct=lambda p: f"{p:.1f}%", startangle=90, counterclock=False,
        textprops={"fontsize": 8}, wedgeprops={"edgecolor": "white", "linewidth": 1})
    for t in autotexts:
        t.set_fontsize(7)
    ax.set_aspect("equal")
    ax.set_title(f"{title}  (n={total})", fontsize=10)


# ── by-Channel DefectMap ──────────────────────────────────

def defect_color(code):
    if code in DEFECT_COLORS:
        return DEFECT_COLORS[code]
    # 未知缺陷码: 用 tab20 均匀上色, 返回 hex 字符串 (openpyxl 需要)
    h = abs(hash(code)) % 20
    rgba = _DEFECT_CMAP(h)
    return "#{:02X}{:02X}{:02X}".format(int(rgba[0] * 255),
                                       int(rgba[1] * 255),
                                       int(rgba[2] * 255))


def compute_channel_stats(judged, wafer_filter=None):
    """统计每个通道 (S1, H1..H32, S2)  across chips 的 OK/失败分布。

    X+/X- (Chip 级 R 均值越限) 的 Chip 整片排除 (含 S1/S2 通道)。
    返回 (stats, x_ng):
      stats: {Hd: {"total": n, "ok": ok, "fail": fail, "dominant": code,
                   "codes": {code: count}, "yield": yield_pct}}
      x_ng:  被 X+/X- 排除的 Chip 数
    """
    df = judged
    if wafer_filter and wafer_filter != "ALL":
        df = df.filter(pl.col("WaferID") == wafer_filter)
    # 排除 X+/X- Reject 的 Chip (整 Chip 数据, 含 S1/S2)
    x_chips = df.filter(pl.col("CODE").is_in(["X+", "X-"])
                        ).select(["WaferID", "ChipID"]).unique()
    x_ng = x_chips.height
    if x_ng > 0:
        df = df.join(x_chips, on=["WaferID", "ChipID"], how="anti")
    stats = {}
    for hd in CHANNEL_ORDER:
        sub = df.filter(pl.col("Hd") == hd)
        total = sub.height
        if total == 0:
            stats[hd] = {"total": 0, "ok": 0, "fail": 0, "dominant": None,
                         "codes": {}, "yield": None}
            continue
        ok = sub.filter(pl.col("CODE") == "OK").height
        fail = total - ok
        code_cnt = {}
        for code in sub.filter(pl.col("CODE") != "OK")["CODE"].to_list():
            code_cnt[code] = code_cnt.get(code, 0) + 1
        dominant = max(code_cnt, key=code_cnt.get) if code_cnt else "OK"
        stats[hd] = {
            "total": total, "ok": ok, "fail": fail,
            "dominant": dominant, "codes": code_cnt,
            "yield": ok / total * 100,
        }
    return stats, x_ng


# CHANNEL_ORDER 中每个通道的列索引 (S1=0, H1=1, ..., H32=32, S2=33)
_CHAN_IDX = {hd: i for i, hd in enumerate(CHANNEL_ORDER)}
_NCHAN = len(CHANNEL_ORDER)


def generate_channel_grid(judged, wafer_id=None):
    """构建 by-Channel DefectMap 网格: 行=ChipRow, 列=(ChipCol组 × 34通道)。

    每个 cell 存储该 (ChipRow, ChipCol, Channel) 的缺陷码。
    列布局: 每个 ChipCol 占 34 列 (S1, H1..H32, S2), 多个 ChipCol 横向排列。
    wafer_id=None 时使用全部数据。
    """
    sub = judged
    if wafer_id is not None:
        sub = judged.filter(pl.col("WaferID") == wafer_id)
    # ChipCol 排序映射到 0..n-1
    chip_cols = sorted(set(
        c for _, c in (parse_chip_rowcol(cid) for cid in sub["ChipID"].to_list())
        if c is not None))
    col_rank = {c: i for i, c in enumerate(chip_cols)}
    grid = {}
    rows, cols = [], []
    for cid, hd, code in zip(sub["ChipID"], sub["Hd"], sub["CODE"]):
        r, c = parse_chip_rowcol(cid)
        if r is None or c is None or hd not in _CHAN_IDX:
            continue
        ci = col_rank.get(c)
        if ci is None:
            continue
        col = ci * _NCHAN + _CHAN_IDX[hd]
        grid[(r, col)] = code
        rows.append(r)
        cols.append(col)
    if not rows:
        return None
    return {
        "grid": grid,
        "r_min": min(rows), "r_max": max(rows),
        "c_min": min(cols), "c_max": max(cols),
        "chip_cols": chip_cols,
    }


def plot_channel_map(ax, grid_data, title="By-Channel Defect Map"):
    """绘制 by-Channel DefectMap: 行=ChipRow, 列=通道位置 (S1|H1..H32|S2 per ChipCol)。

    cell 颜色 = 该 (chip, channel) 的缺陷码, OK 为绿色。
    仅保留 ChipCol 组间与 S1/H1-H32/S2 组分隔线, 去除逐 cell 黑线。
    """
    if grid_data is None:
        ax.text(0.5, 0.5, "No Data", ha="center", va="center", fontsize=14)
        ax.set_title(title)
        return
    grid = grid_data["grid"]
    r_min, r_max = grid_data["r_min"], grid_data["r_max"]
    c_min, c_max = grid_data["c_min"], grid_data["c_max"]
    n_rows = r_max - r_min + 1
    n_cols = c_max - c_min + 1
    chip_cols = grid_data.get("chip_cols", [])
    for r in range(r_min, r_max + 1):
        for c in range(c_min, c_max + 1):
            code = grid.get((r, c))
            if code is None:
                color = "#E0E0E0"
            else:
                color = defect_color(code) if code != "OK" else DEFECT_COLORS["OK"]
            i = r - r_min
            j = c - c_min
            # 去除逐 cell 黑线, 仅靠颜色区分
            ax.add_patch(plt.Rectangle((j, n_rows - 1 - i), 1, 1,
                                       facecolor=color, edgecolor="none",
                                       linewidth=0))
    ax.set_xlim(0, n_cols)
    ax.set_ylim(0, n_rows)
    ax.set_aspect("auto")
    ax.set_box_aspect(1)
    ax.set_xlabel("Map_Col (S1 | H1..H32 | S2  per ChipCol)")
    ax.set_ylabel("ChipRow")
    ax.set_title(title, fontsize=11)
    # X 轴刻度: 每组中心标 ChipCol
    if chip_cols:
        centers = [c_min + (i + 0.5) * _NCHAN for i in range(len(chip_cols))]
        ax.set_xticks(centers)
        ax.set_xticklabels([str(c) for c in chip_cols], fontsize=7)
    ax.tick_params(labelsize=7)


def plot_channel_defect_pie(ax, channel_stats, title="By-Channel Defect"):
    """by-Channel DefectMap 下方的通道缺陷分布饼图。

    统计每个通道 (S1,H1..H32,S2) 在其全部 chip 上的主导缺陷码,
    按主导码聚合通道数绘制饼图。OK 通道与无数据通道不参与统计与显示,
    饼图仅反映发生失败的通道的缺陷构成。
    X+ / X- (Chip 级 R 均值越限) 不参与统计: 计算各通道主导码时先剔除
    X+/X- 记录, 再在剩余缺陷码中取主导; 若剔除后该通道无缺陷则不计入。
    """
    _X_CODES = ("X+", "X-")
    code_cnt = {}
    has_data = False
    for hd in CHANNEL_ORDER:
        s = channel_stats.get(hd)
        if not s or s.get("total", 0) == 0:
            continue
        has_data = True
        # 先排除 X+/X- 再取主导码, 避免 X Chip 把整片 HD 通道计入饼图
        codes = {c: n for c, n in (s.get("codes") or {}).items()
                 if c not in _X_CODES}
        if not codes:
            continue
        dom = max(codes, key=codes.get)
        code_cnt[dom] = code_cnt.get(dom, 0) + 1
    if not has_data:
        ax.text(0.5, 0.5, "No Data", ha="center", va="center", fontsize=11)
        ax.set_title(title, fontsize=10)
        ax.axis("off")
        return
    if not code_cnt:
        ax.text(0.5, 0.5, "All Channels OK\n(X+/- excluded)", ha="center",
                va="center", fontsize=10, color=DEFECT_COLORS["OK"])
        ax.set_title(title, fontsize=10)
        ax.axis("off")
        return
    # 按通道数降序
    items = sorted(code_cnt.items(), key=lambda kv: -kv[1])
    labels = [c for c, _ in items]
    sizes = [n for _, n in items]
    colors = [defect_color(c) for c, _ in items]
    total = sum(sizes)
    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors,
        autopct=lambda p: f"{p:.1f}%", startangle=90, counterclock=False,
        textprops={"fontsize": 8}, wedgeprops={"edgecolor": "white", "linewidth": 1})
    for t in autotexts:
        t.set_fontsize(7)
    ax.set_aspect("equal")
    ax.set_title(f"{title}  (n={total} NG channels)", fontsize=10)


# ── Excel 导出 ─────────────────────────────────────────────

def _col_letter(idx):
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _kde_xy(values, lsl=None, usl=None, n=200):
    """高斯 KDE 密度曲线, 返回 (xs, ys, ymax)。数据不足返回 None。

    视图窗口按 0.5%~99.5% 分位裁剪 (剔除极端离群点, 否则坐标轴被拉得过宽),
    密度值仍基于全部样本计算。
    """
    v = np.asarray([x for x in values if x is not None], dtype=float)
    v = v[np.isfinite(v)]
    if v.size < 2:
        return None
    if v.size >= 20:
        q_lo, q_hi = np.percentile(v, [0.5, 99.5])
        lo, hi = float(q_lo), float(q_hi)
        core = v[(v >= q_lo) & (v <= q_hi)]
    else:
        lo, hi = float(v.min()), float(v.max())
        core = v
    if hi <= lo:
        hi = lo + 1.0
    pad = (hi - lo) * 0.15
    lo -= pad
    hi += pad
    # 确保 Spec 限包含在视图范围内
    span = hi - lo
    if lsl is not None:
        lo = min(lo, lsl - span * 0.03)
    if usl is not None:
        hi = max(hi, usl + span * 0.03)
    xs = np.linspace(lo, hi, n)
    # 带宽用裁剪后的样本估计, 避免被离群点撑大而过度平滑
    sd = float(core.std())
    bw = 1.06 * sd * core.size ** (-1 / 5) if sd > 0 else (hi - lo) / 20
    if bw <= 0:
        bw = (hi - lo) / 20
    z = (xs[:, None] - v[None, :]) / bw
    dens = np.exp(-0.5 * z ** 2).sum(axis=1) / (v.size * bw * np.sqrt(2 * np.pi))
    return xs, dens, float(dens.max())


def _trim_stats(values):
    """3σ 剔除异常后的统计量: (mean, sigma, cv%, n_keep, n_total)。

    与 compute_dynamic_limits_32hd 口径一致: 按样本标准差 (ddof=1) 单次 3σ 剔除。
    样本不足返回 None。
    """
    v = np.asarray([x for x in values if x is not None], dtype=float)
    v = v[np.isfinite(v)]
    n_total = int(v.size)
    if v.size < 2:
        return None
    m, s = float(v.mean()), float(v.std(ddof=1))
    if s > 0:
        v = v[np.abs(v - m) <= 3 * s]
    if v.size < 2:
        return None
    m, s = float(v.mean()), float(v.std(ddof=1))
    cv = (s / abs(m) * 100.0) if m != 0 else None
    return m, s, cv, int(v.size), n_total


def _fmt_sigma_stats(stats):
    """_trim_stats 结果 -> 图表标注文本。"""
    if stats is None:
        return "3σ剔除后: 样本不足"
    m, s, cv, n_keep, n_total = stats
    cv_txt = f"{cv:.2f}%" if cv is not None else "N/A"
    return (f"3σ剔除后: Mean={m:.4g}  Sigma={s:.4g}  CV%={cv_txt}"
            f"  (N={n_keep}/{n_total})")


def _chart_title(line1, line2):
    """两行图表标题: 第一行加粗, 第二行为统计标注。"""
    return Title(tx=Text(rich=RichText(p=[
        Paragraph(pPr=ParagraphProperties(
            defRPr=CharacterProperties(sz=1200, b=True)),
            r=[RegularTextRun(t=line1)]),
        Paragraph(pPr=ParagraphProperties(
            defRPr=CharacterProperties(sz=900)),
            r=[RegularTextRun(t=line2)]),
    ])))


def _nice_axis_range(lo, hi):
    """把显式坐标轴范围对齐到整齐刻度。

    否则 Excel 会直接按 min 值起刻度并以全精度显示标签
    (如 75.914826), 图面很难看。按 span/5 取 10 的整数次幂为步长,
    向外取整保证数据不被裁掉。
    """
    span = hi - lo
    if not np.isfinite(span) or span <= 0:
        return lo, hi
    step = 10.0 ** np.floor(np.log10(span / 5.0))
    n_lo = float(np.floor(lo / step) * step)
    n_hi = float(np.ceil(hi / step) * step)
    if n_hi <= n_lo:
        n_hi = n_lo + step
    return n_lo, n_hi


def write_excel(result, save_path):
    wb = Workbook()
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4",
                              fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    grade_fill = {"A": PatternFill(start_color="C6EFCE", fill_type="solid"),
                  "B": PatternFill(start_color="FFEB9C", fill_type="solid"),
                  "NG": PatternFill(start_color="FFC7CE", fill_type="solid")}

    chips = result["chips"]
    wafers = [r["WaferID"] for r in result["wafer_summary"].rows(named=True)]

    # ── Sheet 1: Grade Summary ──
    ws = wb.active
    ws.title = "Grade Summary"
    meta = [("Source", result["source"]),
            ("Analyzed At", result["analyzed_at"]),
            ("Files", len(result["files"]))]
    row = 1
    for k, v in meta:
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v))
        row += 1
    row += 1
    headers = ["Wafer ID", "Total", "A", "B", "NG", "A%", "B%", "Pass%(A+B)"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    row += 1
    for r in result["wafer_summary"].rows(named=True):
        vals = [r["WaferID"], r["Total"], r["A"], r["B"], r["NG"],
                f"{r['A%']:.2f}", f"{r['B%']:.2f}", f"{r['Pass%']:.2f}"]
        for c, v in enumerate(vals, 1):
            ws.cell(row=row, column=c, value=v).border = thin
            ws.cell(row=row, column=c).alignment = center
        row += 1
    tot = result["wafer_summary"]
    t_total = tot["Total"].sum()
    t_a, t_b, t_ng = tot["A"].sum(), tot["B"].sum(), tot["NG"].sum()
    tvals = ["TOTAL", t_total, t_a, t_b, t_ng,
             f"{t_a / t_total * 100:.2f}" if t_total else "0",
             f"{t_b / t_total * 100:.2f}" if t_total else "0",
             f"{(t_a + t_b) / t_total * 100:.2f}" if t_total else "0"]
    for c, v in enumerate(tvals, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = Font(bold=True)
        cell.alignment = center
        cell.border = thin
    row += 2

    # NG 根因分布
    ws.cell(row=row, column=1, value="NG Root Cause").font = Font(bold=True, size=12)
    row += 1
    for c, h in enumerate(["Wafer ID", "Root Code", "Count"], 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    row += 1
    for r in result["cause"].rows(named=True):
        for c, v in enumerate([r["WaferID"], r["ROOT_CODE"], r["CNT"]], 1):
            ws.cell(row=row, column=c, value=v).border = thin
            ws.cell(row=row, column=c).alignment = center
        row += 1
    for i in range(1, 9):
        ws.column_dimensions[_col_letter(i)].width = 13

    # ── Sheet 2: Chip Detail ──
    ws2 = wb.create_sheet("Chip Detail")
    chan_cols = CHANNEL_ORDER
    headers2 = ["WaferID", "ChipID", "Grade", "HD_OK(/32)", "HD_N",
                "S1S2", "Root Code", "Root Ch"] + chan_cols
    for c, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    row = 2
    for r in chips.sort(["WaferID", "ChipID"]).rows(named=True):
        s1 = r.get("S1", None)
        s2 = r.get("S2", None)
        ss = "OK" if s1 == "OK" and s2 == "OK" else "NG"
        root = r.get("ROOT_CODE")
        if r["GRADE"] != "NG":
            root = "-"
        root_ch = r.get("ROOT_CH") if r["GRADE"] == "NG" else "-"
        vals = [r["WaferID"], r["ChipID"], r["GRADE"], r["HD_OK"], r["HD_N"],
                ss, root, root_ch] + [r.get(c) for c in chan_cols]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(row=row, column=c, value=v)
            cell.border = thin
            cell.alignment = center
        gcell = ws2.cell(row=row, column=3)
        gcell.fill = grade_fill.get(r["GRADE"], grade_fill["NG"])
        row += 1
    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 12
    for i in range(3, len(headers2) + 1):
        ws2.column_dimensions[_col_letter(i)].width = 7
    ws2.freeze_panes = "C2"

    # ── Sheet 3: Spec Used ──
    ws3 = wb.create_sheet("Spec Used")
    headers3 = ["Group", "ColumnName", "ShortName", "RejectPriority",
                "LSL", "USL", "LSL_Code", "USL_Code", "参与判定"]
    for c, h in enumerate(headers3, 1):
        cell = ws3.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    row = 2
    bj = result.get("bj_limits", {}).get("32HD", {})
    dyn_shortnames = {"R", "MR_Ratio"}
    for g in ("32HD", "S1S2"):
        for s in result["spec_groups"][g]:
            used = "Y" if s["Priority"] is not None else "N(优先级空)"
            lsl, usl = s["LSL"], s["USL"]
            note = ""
            if g == "32HD" and s["ShortName"] in dyn_shortnames:
                # 实际使用各 Wafer 的动态统计限, DefectCode LSL/USL 仅为均值规格
                note = "动态(见下)"
                lsl = usl = "见下方"
            elif s["Priority"] is not None and lsl is None and usl is None:
                el = bj.get(s["ColumnName"]) or bj.get(s["ShortName"])
                if el:
                    lsl, usl = el[0], el[1]
                    note = "Y(手动)"
                else:
                    used = "N(无Spec)"
            vals = [g, s["ColumnName"], s["ShortName"], s["Priority"],
                    lsl, usl, s["LSL_Code"], s["USL_Code"],
                    note if note else used]
            for c, v in enumerate(vals, 1):
                ws3.cell(row=row, column=c, value=v).border = thin
        row += 1
    # 动态统计限明细 (按 Chip)
    row += 2
    ws3.cell(row=row, column=1, value="32HD 动态统计限 (按 Chip; 3σ剔除后均值 ±波段)").font = Font(bold=True, size=12)
    row += 1
    dyn_headers = ["WaferID", "ChipID", "ShortName", "过滤后均值", "LSL", "USL",
                   "有效样本", "总样本"]
    for c, h in enumerate(dyn_headers, 1):
        cell = ws3.cell(row=row, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    row += 1
    dyn = result.get("dynamic_limits")
    if dyn is not None and dyn.height > 0:
        target = {s["ColumnName"]: s["ShortName"]
                  for s in result["spec_groups"]["32HD"]
                  if s["ShortName"] in ("R", "MR_Ratio")}
        for r in dyn.rows(named=True):
            for col, short in target.items():
                lsl_c, usl_c = f"{col}__LSL", f"{col}__USL"
                if lsl_c not in dyn.columns:
                    continue
                lsl = r.get(lsl_c)
                if lsl is None:
                    continue
                vals = [r["WaferID"], r["ChipID"], short,
                        round(r.get(f"{col}__mean"), 4),
                        round(lsl, 4), round(r.get(usl_c), 4),
                        r.get(f"{col}__keep_n"), r.get(f"{col}__n")]
                for c, v in enumerate(vals, 1):
                    ws3.cell(row=row, column=c, value=v).border = thin
                row += 1
    for i in range(1, len(headers3) + 1):
        ws3.column_dimensions[_col_letter(i)].width = 15

    # ── Sheet 4: Channel Yield (by-Channel) ──
    ws4 = wb.create_sheet("Channel Yield")
    ch_headers = ["WaferID", "Hd", "Total", "OK", "Fail", "Yield%",
                  "DominantDefect", "X_NG_Chips"]
    for c, h in enumerate(ch_headers, 1):
        cell = ws4.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    row = 2
    wafers = [r["WaferID"] for r in result["wafer_summary"].rows(named=True)]
    for wid in wafers:
        ch_stats, x_ng = compute_channel_stats(result["judged"],
                                               wafer_filter=wid)
        if x_ng > 0:
            c = ws4.cell(row=row, column=1,
                         value=f"Wafer: {wid}  X NG Chips (R均值越限, 已排除): {x_ng}")
            c.font = Font(bold=True, color="CC0000")
            row += 1
        for hd in CHANNEL_ORDER:
            s = ch_stats.get(hd, {})
            vals = [wid, hd, s.get("total", 0), s.get("ok", 0),
                    s.get("fail", 0), round(s.get("yield") or 0, 2),
                    s.get("dominant", "-"), x_ng]
            for c, v in enumerate(vals, 1):
                cell = ws4.cell(row=row, column=c, value=v)
                cell.border = thin
                cell.alignment = center
                # 良率 < 100 标红
                if c == 6 and s.get("total", 0) > 0 and s.get("fail", 0) > 0:
                    cell.fill = PatternFill(start_color="FFC7CE", fill_type="solid")
            row += 1
    for i in range(1, len(ch_headers) + 1):
        ws4.column_dimensions[_col_letter(i)].width = 14

    # ── Sheet 5: Chip DefectMap ──
    ws5 = wb.create_sheet("Chip DefectMap")
    _defect_fill_cache = {}
    def _fill(hex_color):
        if hex_color not in _defect_fill_cache:
            _defect_fill_cache[hex_color] = PatternFill(
                start_color=hex_color, end_color=hex_color, fill_type="solid")
        return _defect_fill_cache[hex_color]
    # grade color hex
    grade_hex = {"A": "C6EFCE", "B": "FFEB9C", "NG": "FFC7CE"}
    untested_hex = "E0E0E0"

    cur_row = 1
    for wid in wafers:
        grid_data = generate_chip_grid(chips, wid)
        ws5.cell(row=cur_row, column=1, value=f"Wafer: {wid}").font = Font(bold=True, size=12)
        cur_row += 1
        if grid_data is None:
            ws5.cell(row=cur_row, column=1, value="(no data)")
            cur_row += 3
            continue
        grid = grid_data["grid"]
        r_min, r_max = grid_data["r_min"], grid_data["r_max"]
        c_min, c_max = grid_data["c_min"], grid_data["c_max"]
        # 列头: ChipCol
        for j, c in enumerate(range(c_min, c_max + 1)):
            cell = ws5.cell(row=cur_row, column=2 + j, value=c)
            cell.font = Font(bold=True)
            cell.alignment = center
        # 行: ChipRow
        for i, r in enumerate(range(r_min, r_max + 1)):
            ws5.cell(row=cur_row + 1 + i, column=1, value=r).font = Font(bold=True)
            ws5.cell(row=cur_row + 1 + i, column=1).alignment = center
            for j, c in enumerate(range(c_min, c_max + 1)):
                val = grid.get((r, c))
                cell = ws5.cell(row=cur_row + 1 + i, column=2 + j, value=val or "")
                cell.alignment = center
                cell.fill = _fill(grade_hex.get(val, untested_hex))
                cell.border = thin
        cur_row += (r_max - r_min + 1) + 3
    ws5.column_dimensions["A"].width = 8
    for j in range(1, 60):
        ws5.column_dimensions[_col_letter(j + 1)].width = 5

    # ── Sheet 6: Channel DefectMap ──
    ws6 = wb.create_sheet("Channel DefectMap")
    # defect code -> hex (去掉 #)
    dc_hex = {k: v.lstrip("#") for k, v in DEFECT_COLORS.items()}
    ok_hex = DEFECT_COLORS["OK"].lstrip("#")

    cur_row = 1
    for wid in wafers:
        grid_data = generate_channel_grid(result["judged"], wafer_id=wid)
        ch_stats, x_ng = compute_channel_stats(result["judged"],
                                               wafer_filter=wid)
        total_ok = sum(s["ok"] for s in ch_stats.values())
        total_n = sum(s["total"] for s in ch_stats.values())
        yld = total_ok / total_n * 100 if total_n else 0
        x_note = f"  X NG Chips: {x_ng}" if x_ng > 0 else ""
        ws6.cell(row=cur_row, column=1,
                 value=f"Wafer: {wid}  Channel Yield: {yld:.1f}%{x_note}"
                 ).font = Font(bold=True, size=12)
        cur_row += 1
        if grid_data is None:
            ws6.cell(row=cur_row, column=1, value="(no data)")
            cur_row += 3
            continue
        grid = grid_data["grid"]
        r_min, r_max = grid_data["r_min"], grid_data["r_max"]
        c_min, c_max = grid_data["c_min"], grid_data["c_max"]
        chip_cols = grid_data.get("chip_cols", [])
        # 列头: 每个 ChipCol 组标一次 (跨 34 列合并)
        for ci, cc in enumerate(chip_cols):
            start_col = 2 + ci * _NCHAN
            ws6.cell(row=cur_row, column=start_col, value=f"Col{cc}\n(S1|H1..H32|S2)")
            ws6.cell(row=cur_row, column=start_col).font = Font(bold=True)
            ws6.cell(row=cur_row, column=start_col).alignment = center
            if _NCHAN > 1:
                ws6.merge_cells(start_row=cur_row, start_column=start_col,
                                end_row=cur_row, end_column=start_col + _NCHAN - 1)
        # 行: ChipRow
        for i, r in enumerate(range(r_min, r_max + 1)):
            ws6.cell(row=cur_row + 1 + i, column=1, value=r).font = Font(bold=True)
            ws6.cell(row=cur_row + 1 + i, column=1).alignment = center
            for j, c in enumerate(range(c_min, c_max + 1)):
                code = grid.get((r, c))
                cell = ws6.cell(row=cur_row + 1 + i, column=2 + j,
                                value="" if code is None else code)
                cell.alignment = center
                if code is None:
                    cell.fill = _fill(untested_hex)
                elif code == "OK":
                    cell.fill = _fill(ok_hex)
                else:
                    cell.fill = _fill(dc_hex.get(code, defect_color(code).lstrip("#")))
        cur_row += (r_max - r_min + 1) + 3
    ws6.column_dimensions["A"].width = 8
    for j in range(1, 400):
        ws6.column_dimensions[_col_letter(j + 1)].width = 3

    # ── Sheet 7: Density (R / MR_Ratio / ASYM / BJMP2% 密度分布) ──
    ws7 = wb.create_sheet("Density")
    ws7d = wb.create_sheet("_density_data")   # 图表数据 (隐藏)
    ws7d.sheet_state = "hidden"
    # 目标参数: 请求的 ShortName (宽松匹配) -> spec
    import difflib
    specs32 = result["spec_groups"]["32HD"]
    specsss = result["spec_groups"]["S1S2"]

    def _match_spec(req, specs):
        names = {s["ShortName"]: s for s in specs}
        if req in names:
            return names[req]
        norm = {k.upper().replace("%", "").replace("_", ""): s
                for k, s in names.items()}
        key = req.upper().replace("%", "").replace("_", "")
        if key in norm:
            return norm[key]
        close = difflib.get_close_matches(key, list(norm.keys()), n=1, cutoff=0.5)
        return norm[close[0]] if close else None

    judged = result["judged"]
    hd_vals = judged.filter(pl.col("Hd").is_in(HD_CHANNELS))
    ss_vals = judged.filter(pl.col("Hd").is_in(SS_CHANNELS))
    # 密度图与统计需先排除 X+/X- (Chip 级 R 均值超规格) 的整 Chip 数据
    dyn = result.get("dynamic_limits")
    _r_spec = next((s for s in specs32 if s["ShortName"] == "R"), None)
    _x_col = f"{_r_spec['ColumnName']}__X_code" if _r_spec else None
    if dyn is not None and _x_col and _x_col in dyn.columns:
        x_df = dyn.filter(pl.col(_x_col).is_not_null()
                         ).select(["WaferID", "ChipID"])
        if x_df.height > 0:
            hd_vals = hd_vals.join(x_df, on=["WaferID", "ChipID"], how="anti")
            ss_vals = ss_vals.join(x_df, on=["WaferID", "ChipID"], how="anti")
    targets = ["R", "MR_Ratio", "ASYM", "BJMP2%"]
    # 布局: 行=参数, 列=分组 (左 32HD / 右 S1S2), 每图占 21 行
    groups = [("32HD", specs32, hd_vals, "A"),
              ("S1S2", specsss, ss_vals, "J")]
    dcol = 1
    for gi, (gname, gspecs, gvals, col_letter) in enumerate(groups):
        for pi, req in enumerate(targets):
            anchor = f"{col_letter}{1 + pi * 21}"
            s = _match_spec(req, gspecs)
            col = s["ColumnName"] if s else None
            title = s["ShortName"] if s else req
            if col is None or col not in gvals.columns:
                ws7.cell(row=1 + pi * 21, column=1 + gi * 9,
                         value=f"({gname} 参数 {req} 无数据)")
                continue
            vals = gvals[col].drop_nulls().to_list()
            lsl, usl = s["LSL"], s["USL"]
            kde = _kde_xy(vals, lsl=lsl, usl=usl)
            if kde is None:
                ws7.cell(row=1 + pi * 21, column=1 + gi * 9,
                         value=f"({gname} 参数 {req} 数据不足)")
                continue
            xs, ys, ymax = kde
            n = len(xs)
            prefix = f"{gname}_{title}"
            # 写数据: x, density, lsl_x, lsl_y, usl_x, usl_y
            ws7d.cell(row=1, column=dcol, value=f"{prefix}_x")
            ws7d.cell(row=1, column=dcol + 1, value=f"{prefix}_y")
            for i in range(n):
                ws7d.cell(row=2 + i, column=dcol, value=float(xs[i]))
                ws7d.cell(row=2 + i, column=dcol + 1, value=float(ys[i]))
            # LSL/USL 竖线 (两点)
            ws7d.cell(row=1, column=dcol + 2, value=f"{prefix}_lsl_x")
            ws7d.cell(row=1, column=dcol + 3, value=f"{prefix}_lsl_y")
            ws7d.cell(row=1, column=dcol + 4, value=f"{prefix}_usl_x")
            ws7d.cell(row=1, column=dcol + 5, value=f"{prefix}_usl_y")
            if lsl is not None:
                ws7d.cell(row=2, column=dcol + 2, value=float(lsl))
                ws7d.cell(row=3, column=dcol + 2, value=float(lsl))
                ws7d.cell(row=2, column=dcol + 3, value=0.0)
                ws7d.cell(row=3, column=dcol + 3, value=float(ymax * 1.1))
            if usl is not None:
                ws7d.cell(row=2, column=dcol + 4, value=float(usl))
                ws7d.cell(row=3, column=dcol + 4, value=float(usl))
                ws7d.cell(row=2, column=dcol + 5, value=0.0)
                ws7d.cell(row=3, column=dcol + 5, value=float(ymax * 1.1))

            chart = ScatterChart()
            chart.title = _chart_title(
                f"{gname} {title} Density (LSL={lsl}, USL={usl})",
                _fmt_sigma_stats(_trim_stats(vals)))
            chart.style = 2
            chart.height = 9.5
            chart.width = 16
            chart.x_axis.title = title
            chart.y_axis.title = "Density"
            chart.scatterStyle = "line"
            chart.legend = None
            # openpyxl ScatterChart 默认两轴 axPos 均为 'l', Excel 会错位/互换;
            # 必须显式指定: x 轴在底部, y 轴在左侧
            chart.x_axis.axPos = "b"
            chart.y_axis.axPos = "l"
            # 显式固定坐标轴范围: 32HD 与 S1S2 各自按本组数据范围设定, 互不相同
            # 并对齐到整齐刻度, 避免 Excel 显示长小数标签
            x_lo, x_hi = _nice_axis_range(float(xs[0]), float(xs[-1]))
            y_hi = float(ymax * 1.15)
            y_lo, y_hi = _nice_axis_range(0.0, y_hi)
            chart.x_axis.scaling.min = x_lo
            chart.x_axis.scaling.max = x_hi
            chart.y_axis.scaling.min = y_lo
            chart.y_axis.scaling.max = y_hi
            # 密度曲线
            # 注意: 引用必须从第 2 行(数据首行)开始, 不能含第 1 行表头。
            # 若含表头文本单元格, Excel 会把它当作首个数据点(值=0),
            # 并把整段 x 引用按文本处理, 导致曲线 x 映射错位、2 点竖线退化不可见。
            xref = Reference(ws7d, min_col=dcol, min_row=2, max_row=1 + n)
            yref = Reference(ws7d, min_col=dcol + 1, min_row=2, max_row=1 + n)
            ser = Series(yref, xref, title="density")
            ser.marker = Marker(symbol="none")
            ser.smooth = True
            ser.graphicalProperties.line.solidFill = "4472C4"
            ser.graphicalProperties.line.width = 22000
            chart.series.append(ser)
            # LSL / USL 红色虚线
            for off in (2, 4):
                lim = lsl if off == 2 else usl
                if lim is None:
                    continue
                # 同样从第 2 行开始, 排除表头, 保证 2 点竖线能正确绘制
                lx = Reference(ws7d, min_col=dcol + off, min_row=2, max_row=3)
                ly = Reference(ws7d, min_col=dcol + off + 1, min_row=2, max_row=3)
                lname = "LSL" if off == 2 else "USL"
                lser = Series(ly, lx, title=lname)
                lser.marker = Marker(symbol="none")
                lser.smooth = False
                lser.graphicalProperties.line.solidFill = "FF0000"
                lser.graphicalProperties.line.dashStyle = "dash"
                lser.graphicalProperties.line.width = 22000
                chart.series.append(lser)
            ws7.add_chart(chart, anchor)
            dcol += 7

    # ── Sheet 8: Date Trend (按测试日期 良率 + 各缺陷率) ──
    ws8 = wb.create_sheet("Date Trend")
    # 每 Chip 的测试时间 (取该 Chip 首条记录的 TestTime)
    chip_tt = (judged.group_by(["WaferID", "ChipID"])
               .agg(pl.col("TestTime").first().alias("TestTime")))
    chip_tt = chip_tt.with_columns(
        pl.col("TestTime").str.to_datetime(strict=False, exact=False)
        .dt.strftime("%Y-%m-%d").alias("Date"))
    cj = chips.join(chip_tt.select(["WaferID", "ChipID", "Date"]),
                    on=["WaferID", "ChipID"], how="left")
    cj = cj.filter(pl.col("Date").is_not_null() & (pl.col("Date") != ""))
    if cj.height == 0:
        ws8.cell(row=1, column=1, value="(无测试日期数据)")
    else:
        # 每日期 良率
        by_date = (cj.group_by("Date").agg(
            pl.len().alias("Total"),
            (pl.col("GRADE") == "A").sum().alias("A"),
            (pl.col("GRADE") == "B").sum().alias("B"),
            (pl.col("GRADE") == "NG").sum().alias("NG"),
        ).sort("Date"))
        by_date = by_date.with_columns(
            ((pl.col("A") + pl.col("B")) / pl.col("Total") * 100)
            .round(2).alias("Yield%"))
        # 每日期 各缺陷码 (NG 根因) 缺陷率
        def_by_date = (cj.filter(pl.col("GRADE") == "NG")
                       .group_by(["Date", "ROOT_CODE"]).agg(pl.len().alias("N")))
        code_totals = (def_by_date.group_by("ROOT_CODE")
                       .agg(pl.col("N").sum().alias("T"))
                       .sort("T", descending=True))
        def_codes = [r["ROOT_CODE"] for r in code_totals.rows(named=True)]
        # pivot: Date x code -> rate
        dates = by_date["Date"].to_list()
        tot_map = {r["Date"]: r["Total"] for r in by_date.rows(named=True)}
        rate = {}
        for r in def_by_date.rows(named=True):
            t = tot_map.get(r["Date"], 0)
            rate[(r["Date"], r["ROOT_CODE"])] = (r["N"] / t * 100) if t else 0
        # 写表: Date | Total | A | B | NG | Yield% | code1% ...
        headers8 = ["Date", "Total", "A", "B", "NG", "Yield%"] + \
                   [f"{c}%" for c in def_codes]
        for c, h in enumerate(headers8, 1):
            cell = ws8.cell(row=1, column=c, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = thin
        for i, r in enumerate(by_date.rows(named=True)):
            row = 2 + i
            d = r["Date"]
            vals = [d, r["Total"], r["A"], r["B"], r["NG"], r["Yield%"]]
            vals += [round(rate.get((d, c), 0.0), 2) for c in def_codes]
            for c, v in enumerate(vals, 1):
                cell = ws8.cell(row=row, column=c, value=v)
                cell.border = thin
                cell.alignment = center
        nd = len(dates)
        # 良率折线图
        lc1 = LineChart()
        lc1.title = "Yield% by Test Date"
        lc1.style = 12
        lc1.height = 8.5
        lc1.width = 18
        lc1.y_axis.title = "Yield%"
        lc1.x_axis.title = "Date"
        ydata = Reference(ws8, min_col=6, min_row=1, max_row=1 + nd)
        cats = Reference(ws8, min_col=1, min_row=2, max_row=1 + nd)
        lc1.add_data(ydata, titles_from_data=True)
        lc1.set_categories(cats)
        for s_ in lc1.series:
            s_.smooth = False
            s_.marker = Marker(symbol="circle", size=6)
            s_.graphicalProperties.line.solidFill = "2E7D32"
            s_.graphicalProperties.line.width = 22000
        ws8.add_chart(lc1, f"A{nd + 4}")
        # 各缺陷率折线图
        if def_codes:
            lc2 = LineChart()
            lc2.title = "Defect Rate% by Test Date"
            lc2.style = 12
            lc2.height = 8.5
            lc2.width = 18
            lc2.y_axis.title = "Defect Rate%"
            lc2.x_axis.title = "Date"
            ddata = Reference(ws8, min_col=7, max_col=6 + len(def_codes),
                              min_row=1, max_row=1 + nd)
            lc2.add_data(ddata, titles_from_data=True)
            lc2.set_categories(cats)
            for idx2, s_ in enumerate(lc2.series):
                s_.smooth = False
                s_.marker = Marker(symbol="circle", size=5)
                col_hex = DEFECT_COLORS.get(def_codes[idx2], "#333333").lstrip("#")
                s_.graphicalProperties.line.solidFill = col_hex
                s_.graphicalProperties.line.width = 20000
            ws8.add_chart(lc2, f"J{nd + 4}")
        for i in range(1, len(headers8) + 1):
            ws8.column_dimensions[_col_letter(i)].width = 12

    wb.save(save_path)


# ── GUI ────────────────────────────────────────────────────

class RHYieldApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RH 良率分析 (Polars)")
        self.root.geometry("1280x860")
        self.root.minsize(1000, 640)

        self.result = None
        self.spec_groups = None
        self.cfg = load_config()

        self._wafer_var = tk.StringVar(value="ALL")
        self._build_ui()
        self._load_saved()

    # ── UI 构建 ──
    def _build_ui(self):
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left = ttk.Frame(main, width=360)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        left.pack_propagate(False)

        right = ttk.Frame(main)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_input_panel(left)
        self._build_spec_panel(left)
        self._build_result_panel(left)
        self._build_chart_panel(right)

    def _build_input_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Input", padding=8)
        frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(frame, text="数据源 (文件夹或文件):").pack(anchor=tk.W)
        src_frame = ttk.Frame(frame)
        src_frame.pack(fill=tk.X, pady=2)
        self.src_var = tk.StringVar()
        ttk.Entry(src_frame, textvariable=self.src_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(src_frame, text="文件夹...", width=8,
                   command=self._browse_folder).pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(src_frame, text="文件...", width=6,
                   command=self._browse_files).pack(side=tk.LEFT, padx=(2, 0))

        ttk.Label(frame, text="DefectCode.xlsx:").pack(anchor=tk.W, pady=(5, 0))
        dc_frame = ttk.Frame(frame)
        dc_frame.pack(fill=tk.X, pady=2)
        self.dc_var = tk.StringVar(value=str(DEFECT_CODE_FILE))
        ttk.Entry(dc_frame, textvariable=self.dc_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(dc_frame, text="...", width=3,
                   command=self._browse_defectcode).pack(side=tk.LEFT, padx=(2, 0))

        self.only_r_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="仅统计 ChipID 含 'R' 的 Chip",
                        variable=self.only_r_var).pack(anchor=tk.W, pady=(5, 0))

        # X+/- 判定用 R 均值静态限 (默认取 DefectCode R 的 LSL/USL)
        x_frame = ttk.Frame(frame)
        x_frame.pack(fill=tk.X, pady=(3, 0))
        ttk.Label(x_frame, text="X+/- R均值 LSL:").pack(side=tk.LEFT)
        self.x_lsl_var = tk.StringVar()
        ttk.Entry(x_frame, textvariable=self.x_lsl_var, width=7).pack(
            side=tk.LEFT, padx=(2, 6))
        ttk.Label(x_frame, text="USL:").pack(side=tk.LEFT)
        self.x_usl_var = tk.StringVar()
        ttk.Entry(x_frame, textvariable=self.x_usl_var, width=7).pack(
            side=tk.LEFT, padx=(2, 0))

        ttk.Label(frame, text="Wafer:").pack(anchor=tk.W, pady=(5, 0))
        self.wafer_combo = ttk.Combobox(frame, textvariable=self._wafer_var,
                                        state="readonly")
        self.wafer_combo.pack(fill=tk.X, pady=2)
        self.wafer_combo.bind("<<ComboboxSelected>>",
                              lambda e: self._display_chart())

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        self.analyze_btn = ttk.Button(btn_frame, text="开始分析", command=self._analyze)
        self.analyze_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(btn_frame, text="导出 Excel", command=self._export).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

    def _build_spec_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Spec (DefectCode)", padding=8)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        cols = ("group", "name", "prio", "lsl", "usl", "codes", "used")
        self.spec_tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        heads = [("group", "组", 45), ("name", "参数", 105), ("prio", "优先级", 45),
                 ("lsl", "LSL", 50), ("usl", "USL", 50), ("codes", "Code", 70),
                 ("used", "判定", 55)]
        for key, text, w in heads:
            self.spec_tree.heading(key, text=text)
            self.spec_tree.column(key, width=w, anchor=tk.CENTER)
        ysb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.spec_tree.yview)
        self.spec_tree.configure(yscrollcommand=ysb.set)
        self.spec_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_result_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Result", padding=8)
        frame.pack(fill=tk.X, pady=(0, 5))
        self.result_text = tk.Text(frame, height=13, width=44,
                                   font=("Consolas", 9), state=tk.DISABLED)
        sb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=sb.set)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_chart_panel(self, parent):
        self.fig = plt.figure(figsize=(18, 12))
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas.draw()
        tb = ttk.Frame(canvas_frame)
        tb.pack(side=tk.BOTTOM, fill=tk.X)
        NavigationToolbar2Tk(self.canvas, tb)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ── 事件 ──
    def _browse_folder(self):
        d = filedialog.askdirectory(title="选择数据文件夹")
        if d:
            self.src_var.set(d)

    def _browse_files(self):
        fs = filedialog.askopenfilenames(
            title="选择 RH 测试数据文件",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if fs:
            self.src_var.set(";".join(fs))

    def _browse_defectcode(self):
        f = filedialog.askopenfilename(title="选择 DefectCode.xlsx",
                                       filetypes=[("Excel", "*.xlsx")])
        if f:
            self.dc_var.set(f)

    def _load_saved(self):
        self.src_var.set(self.cfg.get("source", ""))
        dc = self.cfg.get("defect_code", str(DEFECT_CODE_FILE))
        self.dc_var.set(dc if Path(dc).exists() else str(DEFECT_CODE_FILE))
        self._load_defect_code()

    def _load_defect_code(self):
        path = self.dc_var.get().strip()
        if not path or not Path(path).exists():
            self.spec_groups = None
            return
        try:
            self.spec_groups = load_defect_code(path)
        except Exception as e:
            messagebox.showerror("Error", f"读取 DefectCode 失败:\n{e}")
            self.spec_groups = None
            return
        self._fill_spec_tree()

    def _fill_spec_tree(self):
        self.spec_tree.delete(*self.spec_tree.get_children())
        if not self.spec_groups:
            return
        # 仅显示参与判定的参数 (Priority 非空), 按优先级升序
        rows = []
        for g in ("32HD", "S1S2"):
            for s in active_params(self.spec_groups[g]):
                if g == "32HD" and s["ShortName"] in ("R", "MR_Ratio"):
                    used = "动态统计限"
                else:
                    used = "Y"
                codes = ",".join(x for x in (s["LSL_Code"], s["USL_Code"]) if x)
                rows.append((s["Priority"], g, s, used, codes))
        rows.sort(key=lambda x: x[0])
        for _, g, s, used, codes in rows:
            self.spec_tree.insert("", tk.END, values=(
                g, s["ShortName"],
                s["Priority"],
                s["LSL"] if s["LSL"] is not None else "/",
                s["USL"] if s["USL"] is not None else "/",
                codes or "/", used))
        # 默认 X+/- R 均值限取自 DefectCode R
        r_spec = next((s for s in self.spec_groups["32HD"]
                       if s["ShortName"] == "R"), None)
        if r_spec:
            if r_spec["LSL"] is not None:
                self.x_lsl_var.set(str(r_spec["LSL"]))
            if r_spec["USL"] is not None:
                self.x_usl_var.set(str(r_spec["USL"]))

    def _save_current_config(self):
        self.cfg["source"] = self.src_var.get()
        self.cfg["defect_code"] = self.dc_var.get()
        save_config(self.cfg)

    # ── 主流程 ──
    def _analyze(self):
        if not self.spec_groups:
            self._load_defect_code()
            if not self.spec_groups:
                messagebox.showwarning("Warning", "请指定有效的 DefectCode.xlsx")
                return
        source = self.src_var.get().strip()
        if not source:
            messagebox.showwarning("Warning", "请选择数据源 (文件夹或文件)")
            return

        self._save_current_config()
        self.analyze_btn.configure(state=tk.DISABLED)
        self.root.configure(cursor="watch")
        only_r = self.only_r_var.get()

        # 解析 X+/- R 均值限 (空则使用 DefectCode 默认)
        def _parse_f(s):
            s = s.strip()
            return float(s) if s else None
        try:
            x_lsl = _parse_f(self.x_lsl_var.get())
            x_usl = _parse_f(self.x_usl_var.get())
        except ValueError:
            x_lsl = x_usl = None
        x_r_limits = {"LSL": x_lsl, "USL": x_usl} if (x_lsl is not None or x_usl is not None) else None

        import queue
        self._result_q = queue.Queue()

        def _run():
            try:
                result, err = analyze(
                    source, self.spec_groups,
                    bj_limits={},
                    only_r_chip=only_r,
                    x_r_limits=x_r_limits,
                    progress_cb=lambda i, n: self._progress(i, n))
            except Exception as e:
                import traceback
                self._result_q.put((None, f"{e}\n\n{traceback.format_exc()[-1500:]}"))
                return
            self._result_q.put((result, err))

        import threading
        threading.Thread(target=_run, daemon=True).start()
        self._poll_result()

    def _poll_result(self):
        if getattr(self, "_prog_n", 0):
            self.root.title(
                f"RH 良率分析 (Polars) - 读取中 {self._prog_i}/{self._prog_n}...")
        try:
            result, err = self._result_q.get_nowait()
        except Exception:
            self.root.after(100, self._poll_result)
            return
        self.root.title("RH 良率分析 (Polars)")
        self._analyze_done(result, err)

    def _analyze_done(self, result, err):
        try:
            if err:
                messagebox.showerror("Error", err)
                return
            if result is None:
                return
            self.result = result
            wafers = [r["WaferID"] for r in result["wafer_summary"].rows(named=True)]
            self.wafer_combo["values"] = ["ALL"] + wafers
            self._wafer_var.set("ALL")
            self._display_result()
            self._display_chart()
        finally:
            self.analyze_btn.configure(state=tk.NORMAL)
            self.root.configure(cursor="")

    def _progress(self, i, n):
        self._prog_i, self._prog_n = i, n

    def _display_result(self):
        r = self.result
        ws = r["wafer_summary"]
        lines = [
            f"数据源: {r['source'][:46]}",
            f"文件数: {len(r['files'])}   分析时间: {r['analyzed_at']}",
            "=" * 52,
            f"{'WaferID':<12}{'Total':>6}{'A':>5}{'B':>5}{'NG':>5}{'Pass%':>9}",
            "-" * 52,
        ]
        for x in ws.rows(named=True):
            lines.append(f"{x['WaferID']:<12}{x['Total']:>6}{x['A']:>5}{x['B']:>5}"
                         f"{x['NG']:>5}{x['Pass%']:>8.2f}%")
        t_total = ws["Total"].sum()
        t_a, t_b = ws["A"].sum(), ws["B"].sum()
        t_ng = ws["NG"].sum()
        lines += ["-" * 52,
                  f"{'TOTAL':<12}{t_total:>6}{t_a:>5}{t_b:>5}{t_ng:>5}"
                  f"{(t_a + t_b) / t_total * 100 if t_total else 0:>8.2f}%",
                  "=" * 52,
                  "32HD 动态统计限 (按 Chip; 3σ剔除后均值 ±波段; R ±20%, MR ±35%, MR下限≥4):"]
        dyn = r.get("dynamic_limits")
        if dyn is not None and dyn.height > 0:
            target = {s["ColumnName"]: s["ShortName"]
                      for s in r["spec_groups"]["32HD"]
                      if s["ShortName"] in ("R", "MR_Ratio")}
            for dr in dyn.rows(named=True):
                for col, short in target.items():
                    lsl_c = f"{col}__LSL"
                    if lsl_c not in dyn.columns:
                        continue
                    lsl = dr.get(lsl_c)
                    if lsl is None:
                        continue
                    lines.append(f"  {dr['WaferID']:<8}{dr['ChipID']:<10}{short:<8}"
                                 f"均值={dr.get(f'{col}__mean', 0):>8.3f}  "
                                 f"LSL={lsl:>7.3f} USL={dr.get(f'{col}__USL', 0):>7.3f}  "
                                 f"(样本 {dr.get(f'{col}__keep_n', 0)}/{dr.get(f'{col}__n', 0)})")
        lines.append("=" * 52)
        lines.append("NG 根因分布 (最低 RejectPriority 缺陷码):")
        cause = r["cause"].group_by("ROOT_CODE").agg(
            pl.col("CNT").sum().alias("CNT")).sort("CNT", descending=True)
        for x in cause.rows(named=True):
            lines.append(f"  {x['ROOT_CODE']:<10}: {x['CNT']}")
        # by-Channel 良率 (仅列出有失败的通道; X+/X- Chip 已排除)
        ch, x_ng = compute_channel_stats(r["judged"],
                                         wafer_filter=self._wafer_var.get())
        fail_chs = [(hd, s) for hd, s in ch.items()
                    if s.get("total", 0) > 0 and s.get("fail", 0) > 0]
        if x_ng > 0:
            lines.append(f"X NG Chips (R均值越限, 已排除): {x_ng}")
        if fail_chs:
            lines.append("=" * 52)
            lines.append("by-Channel 失败通道 (良率 / 主导缺陷):")
            for hd, s in fail_chs:
                lines.append(f"  {hd:<5} 良率={s['yield']:>5.1f}%  "
                             f"({s['ok']}/{s['total']})  主导={s['dominant']}")
        if r["errors"]:
            lines.append("\n警告 (跳过的文件):")
            for e in r["errors"][:8]:
                lines.append(f"  {e[:60]}")
        self.result_text.configure(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, "\n".join(lines))
        self.result_text.configure(state=tk.DISABLED)

    def _display_chart(self):
        if not self.result:
            return
        self.fig.clear()
        chips = self.result["chips"]
        judged = self.result["judged"]
        wafers = [x["WaferID"] for x in self.result["wafer_summary"].rows(named=True)]
        selected = self._wafer_var.get()
        show = wafers if selected == "ALL" else [selected]
        n = len(show)
        # 单晶圆时 ncols=1, 使 by-Chip 与 by-Channel 等宽; 多晶圆按 2/3 列排布
        ncols = 1 if n == 1 else (2 if n <= 4 else 3)
        nrows = (n + ncols - 1) // ncols

        # 左: by-Chip Wafer Map (nrows × ncols); 右: by-Channel DefectMap (跨全部行)
        # 两者等宽且 box_aspect=1 (正方形); 底行放对应 Pie Chart
        gs = self.fig.add_gridspec(
            nrows + 1, ncols + 1,
            width_ratios=[1.0] * ncols + [float(ncols)],
            height_ratios=[1.0] * nrows + [0.55],
            hspace=0.45, wspace=0.3)
        wsum = {x["WaferID"]: x for x in self.result["wafer_summary"].rows(named=True)}
        for idx, wid in enumerate(show):
            r, c = divmod(idx, ncols)
            ax = self.fig.add_subplot(gs[r, c])
            grid = generate_chip_grid(chips, wid)
            w = wsum.get(wid, {})
            title = f"{wid}  A:{w.get('A', 0)} B:{w.get('B', 0)} NG:{w.get('NG', 0)} " \
                    f"Pass:{w.get('Pass%', 0):.1f}%"
            plot_grade_map(ax, grid, title)
            add_grade_legend(ax)
        # 隐藏多余的 by-Chip 子图
        for idx in range(len(show), nrows * ncols):
            r, c = divmod(idx, ncols)
            self.fig.add_subplot(gs[r, c]).set_visible(False)

        # by-Channel DefectMap (右侧, 跨全部行)
        ax_ch = self.fig.add_subplot(gs[0:nrows, ncols])
        ch_wid = None if selected == "ALL" else selected
        ch_grid = generate_channel_grid(judged, wafer_id=ch_wid)
        ch_stats, x_ng = compute_channel_stats(judged, wafer_filter=selected)
        total_ok = sum(s["ok"] for s in ch_stats.values())
        total_n = sum(s["total"] for s in ch_stats.values())
        total_yld = total_ok / total_n * 100 if total_n else 0
        x_note = f"  X NG: {x_ng}" if x_ng > 0 else ""
        plot_channel_map(ax_ch, ch_grid,
                         title=f"By-Channel Defect Map ({selected})  "
                               f"Yield: {total_yld:.1f}%{x_note}")

        # 两个 DefectMap 下方: 对应 Pie Chart
        ax_pie_chip = self.fig.add_subplot(gs[nrows, 0:ncols])
        plot_grade_pie(ax_pie_chip, chips, wafer_filter=selected,
                       title="By-Chip Grade")
        ax_pie_ch = self.fig.add_subplot(gs[nrows, ncols])
        plot_channel_defect_pie(ax_pie_ch, ch_stats,
                                title=f"By-Channel Defect ({selected})")
        self.canvas.draw()

    def _export(self):
        if not self.result:
            messagebox.showwarning("Warning", "请先执行分析")
            return
        # 默认文件名含 LotNo + WaferID + 生成时间戳,
        # 如 268082_J4_TR_MD26A0407_20261019_143012.xlsx
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = self.result.get("file_stem")
        if stem:
            default = f"{stem}_{stamp}.xlsx"
        else:
            default = f"RH_Yield_{stamp}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default,
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            self.root.configure(cursor="watch")
            self.root.update()
            write_excel(self.result, path)
            messagebox.showinfo("Success", f"已保存:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.root.configure(cursor="")


def main():
    root = tk.Tk()
    app = RHYieldApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
