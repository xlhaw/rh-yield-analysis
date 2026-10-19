import os
import re
import sys
import json
from pathlib import Path
from datetime import datetime

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

TR_ROOT = r"\\mdtprod02\MDTWPDC\ZJG_Data\UF190B\Sensor\TR02V02\TR02V02_DW-J4"
TW_ROOT = r"\\mdtprod02\MDTWPDC\ZJG_Data\UF190B\Sensor\TW02V02_DW"

INFO_COLS = ["SHOT_ROW", "SHOT_COL", "WIS_ROW", "WIS_COL"]
DATA_COLS = [f"HD{i:02d}" for i in range(1, 33)]
L_COLS = [f"L{i:02d}" for i in range(1, 33)]
SERVO_COLS = ["S1", "S2"]
PTC_COLS = [f"HD{i:02d}_PTC" for i in range(1, 33)]

TR_ELGA_COLS = [f"RA{i}" for i in range(1, 7)]
TR_ELGB_COLS = [f"RB{i}" for i in range(1, 7)]
TW_ELGA_COLS = [f"RA{i}" for i in range(1, 6)]
TW_ELGB_COLS = [f"RB{i}" for i in range(1, 8)]

# Center 6 ELG channels used for ELG defect judgment. If more than 2 of
# these are out of spec, the die is ELG NG.
ELG_CENTER_COLS = ["RB3", "RA2", "RA3", "RA4", "RA5", "RB4"]

COL_MAP = {
    "Field-Y": "SHOT_ROW", "Field-X": "SHOT_COL",
    "Row": "WIS_ROW", "Column": "WIS_COL",
    "S1_R": "S1", "S2_R": "S2",
    "L": "L",
    "R": "R_L",
    "Actual_X": "ACTUAL_X", "Actual_Y": "ACTUAL_Y",
}
for _i in range(1, 33):
    COL_MAP[f"HD{_i}_R"] = f"HD{_i:02d}"
    COL_MAP[f"HD{_i}_PTC"] = f"HD{_i:02d}_PTC"
    # TW EPS-V3 format names PTC columns "HD{i}_PTC_R"
    COL_MAP[f"HD{_i}_PTC_R"] = f"HD{_i:02d}_PTC"
for _i in range(1, 8):
    COL_MAP[f"RA{_i}"] = f"RA{_i}"
    COL_MAP[f"RB{_i}"] = f"RB{_i}"

# When running as a PyInstaller-frozen exe, __file__ points to the
# temp extraction dir (_MEIxxxx). Use the exe's directory instead so
# config and ChipID mapping files are found next to the exe.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "yieldmap_config.json"
CHIPID_FILE = BASE_DIR / "TR&TW AddrID Vs ChipID.xlsx"

DEFECT_COLORS = {
    "OK": "#4CAF50",
    "IND": "#FF1744",
    "DAT": "#F44336",
    "SEV": "#FF9800",
    "PTC": "#9C27B0",
    "ELG": "#2196F3",
}


def _fmt_lim(lim_pair, idx):
    """Format one side of a (ll, ul) limit pair. Returns '-' when the
    pair is None (e.g. cv-only mode where no 3Sigma limits exist)."""
    if lim_pair is None:
        return "-"
    try:
        return f"{lim_pair[idx]:.4f}"
    except (IndexError, TypeError):
        return "-"


def _fmt_scalar(val):
    """Format a single numeric scalar for the Limits sheet. Returns
    '-' when the value is None (e.g. HF_Sig UL when not in hf_sigma
    mode)."""
    if val is None:
        return "-"
    try:
        return f"{float(val):.4f}"
    except (TypeError, ValueError):
        return "-"


_chipid_cache = {}


def load_chipid_map(sensor_type):
    if sensor_type in _chipid_cache:
        return _chipid_cache[sensor_type]
    if not CHIPID_FILE.exists():
        return None
    df = pl.read_excel(CHIPID_FILE, sheet_name=sensor_type, engine="openpyxl")
    df = df.select(INFO_COLS + ["ChipID"])
    _chipid_cache[sensor_type] = df
    return df


def parse_chipid_rowcol(chipid):
    import re
    m = re.match(r"(\d+)[RW](\d+)", str(chipid))
    if m:
        return int(m.group(1)), int(m.group(2))
    return 0, 0


def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def list_lots(sensor_type):
    root = TR_ROOT if sensor_type == "TR" else TW_ROOT
    p = Path(root)
    if not p.exists():
        return []
    return sorted(
        [d.name for d in p.iterdir() if d.is_dir() and not d.name.startswith(".")]
    )


# --- duplicate-file selection ----------------------------------------------
# When a WaferID has multiple candidate files (repeated measurement or copies
# in different sub-directories), two selection modes are supported:
#   * "size"   - pick the file with the largest byte size (default; the
#                longest measurement / most complete data).
#   * "newest" - pick the newest file. The "newest" timestamp prefers the
#                measurement timestamp embedded in the filename (e.g.
#                "..._2026-05-31 14-07-28.csv") because copies in different
#                folders usually share the same filesystem mtime; it falls
#                back to the filesystem mtime when no timestamp is found.
_RE_FILE_TS = re.compile(r"_(\d{4})-(\d{2})-(\d{2})\s+(\d{2})-(\d{2})-(\d{2})")


def _file_timestamp(fpath):
    """Return a comparable timestamp for *fpath*.

    Prefers the measurement timestamp embedded in the filename (format
    ``YYYY-MM-DD HH-MM-SS``); falls back to the filesystem mtime.
    Returns a ``datetime`` (or ``None`` on failure).
    """
    from datetime import datetime
    m = _RE_FILE_TS.search(fpath.name)
    if m:
        try:
            return datetime(*[int(g) for g in m.groups()])
        except ValueError:
            pass
    try:
        return datetime.fromtimestamp(fpath.stat().st_mtime)
    except OSError:
        return None


def _pick_best(files, dup_mode="size"):
    """Return the best file from *files* according to *dup_mode*.

    ``dup_mode``:
      * ``"size"``   - largest byte size wins
      * ``"newest"`` - newest timestamp wins (filename ts preferred)
    Ties (e.g. identical copies in different folders) keep the first
    encountered file.
    """
    if not files:
        return None
    best_f = files[0]
    for f in files[1:]:
        if dup_mode == "newest":
            ts = _file_timestamp(f)
            best_ts = _file_timestamp(best_f)
            # None sorts smaller than any datetime, so newer wins.
            if ts is not None and (best_ts is None or ts > best_ts):
                best_f = f
        else:  # "size"
            try:
                sz = f.stat().st_size
            except OSError:
                sz = 0
            try:
                best_sz = best_f.stat().st_size
            except OSError:
                best_sz = 0
            if sz > best_sz:
                best_f = f
    return best_f


def _largest_per_wafer(files, dup_mode="size"):
    """Group files by WaferID (extracted from filename) and keep the best
    file for each WaferID according to *dup_mode* ("size" or "newest").

    Returns a list of file paths, one per WaferID, sorted by WaferID.
    """
    groups = {}
    for f in files:
        wid = extract_wafer_id(f)
        groups.setdefault(wid, []).append(f)
    out = []
    for wid in sorted(groups.keys()):
        best = _pick_best(groups[wid], dup_mode=dup_mode)
        if best is not None:
            out.append(best)
    return out


# --- filename tag matching -------------------------------------------------
# Token-boundary regexes so a tag like "HD1" cannot match inside "HD10" and
# "RMIN" cannot match inside "RMIN2". A token boundary is any position that
# is not an alphanumeric (letters/digits); this includes '-', '_', '.', and
# the start/end of the filename.
_RE_HD_LR = re.compile(r"(?:^|[^A-Z0-9])HD([12])(?:-[^A-Z0-9]|[^A-Z0-9]-|-)LR(?:$|[^A-Z0-9])", re.IGNORECASE)
_RE_TAG_RMIN = re.compile(r"(?:^|[^A-Z0-9])RMIN(?:$|[^A-Z0-9])", re.IGNORECASE)
# "EPS" as a prefix is followed by either H (-> EPSH) or -Vn (-> EPS-V2 etc.)
# or a non-alphanumeric boundary. This lets EPSH / EPS-V2 / plain EPS match.
_RE_TAG_EPS = re.compile(r"(?:^|[^A-Z0-9])EPS(?:H|$|[^A-Z0-9]|-[A-Z0-9])", re.IGNORECASE)
_RE_TAG_EPSH = re.compile(r"(?:^|[^A-Z0-9])EPSH(?:$|[^A-Z0-9])", re.IGNORECASE)
_RE_TAG_EPS_V = re.compile(r"(?:^|[^A-Z0-9])EPS-V(\d+)(?:$|[^A-Z0-9])", re.IGNORECASE)


def _match_hd_lr_tag(name):
    """Return 'HD1' / 'HD2' if *name* is an HD1-LR or HD2-LR file, else None.

    Uses a token-boundary regex so HD1 cannot accidentally match HD10.
    """
    m = _RE_HD_LR.search(name)
    if not m:
        return None
    return f"HD{m.group(1)}"


def find_epsh_files(lot_no, sensor_type, eps_format=None, dup_mode="size"):
    """Find EPS data CSV files for a lot.

    Only files under the ``MPAD`` directory are considered, for both TR
    and TW lots.

    ``eps_format`` selects the file variant for TW:
      * ``"EPSH"``   - old TW format, filename contains 'EPSH'
      * ``"EPS-V3"`` - new TW format (EPS-V2/EPS-V3 share the same
        structure: no HD columns, resistance read from HD1-LR/HD2-LR);
        matches any ``EPS-V*`` filename
      * ``None``     - auto (matches any 'EPS' file, used by TR)

    TR always uses the EPSH variant; the ``eps_format`` argument is
    ignored for TR.

    ``dup_mode`` selects how to resolve duplicate files per WaferID:
      * ``"size"``   - keep the largest file by byte size (default)
      * ``"newest"`` - keep the newest file (filename timestamp first,
        filesystem mtime as fallback)

    For each WaferID only one file is kept, so duplicate copies in
    different folders or repeated measurements do not double-count a
    wafer.
    """
    root = TR_ROOT if sensor_type == "TR" else TW_ROOT
    lot_dir = Path(root) / lot_no
    mpad_dir = lot_dir / "MPAD"
    if not mpad_dir.exists():
        return []

    if sensor_type == "TR":
        mode = "EPSH"
    elif eps_format == "EPSH":
        mode = "EPSH"
    elif eps_format == "EPS-V3":
        # New V-style format: EPS-V2 and EPS-V3 are structurally
        # identical (no HD columns; resistance read from LR files).
        mode = "EPS-V"
    else:
        mode = "ANY"

    candidates = []
    for f in mpad_dir.rglob("*.csv"):
        name = f.name

        # Must contain an EPS tag (EPSH / EPS-Vn / generic EPS)
        if not _RE_TAG_EPS.search(name):
            continue

        if mode == "EPSH":
            if not _RE_TAG_EPSH.search(name):
                continue
        elif mode == "EPS-V":
            if not _RE_TAG_EPS_V.search(name):
                continue
        # mode == "ANY" -> accept any EPS match already accepted above

        # For TR, only the RMIN measurement file is used; RMAX files are
        # excluded so a wafer is not double-counted by its two test runs.
        if sensor_type == "TR" and not _RE_TAG_RMIN.search(name):
            continue

        candidates.append(f)
    return _largest_per_wafer(candidates, dup_mode=dup_mode)


_RE_WID_RMIN_SUFFIX = re.compile(r"-RMIN$", re.IGNORECASE)


def extract_wafer_id(fpath):
    wid = fpath.name.split("_")[1] if "_" in fpath.name else "UNKNOWN"
    # TR filenames embed the measurement tag in the WaferID slot (e.g.
    # "W28-RMIN"); strip the "-RMIN" suffix so the WaferID is just "W28".
    return _RE_WID_RMIN_SUFFIX.sub("", wid)


def find_lr_files(lot_no, sensor_type, dup_mode="size"):
    """Find HD1-LR and HD2-LR CSV files for TW lots.

    Only files under the ``MPAD`` directory are considered. For each
    WaferID only one file per HD category is kept (according to
    ``dup_mode``: ``"size"`` for largest, ``"newest"`` for newest),
    across all sub-directories.
    """
    if sensor_type != "TW":
        return {}
    root = TW_ROOT
    lot_dir = Path(root) / lot_no
    mpad_dir = lot_dir / "MPAD"
    if not mpad_dir.exists():
        return {}

    # Group by (WaferID, HD_key), keeping only the best file per group.
    # HD_key is determined via a token-boundary regex match on the filename
    # so HD1 cannot accidentally match HD10, HD12, etc.
    groups = {}
    for f in mpad_dir.rglob("*.csv"):
        hd_key = _match_hd_lr_tag(f.name)
        if hd_key is None:
            continue
        wid = extract_wafer_id(f)
        groups.setdefault((wid, hd_key), []).append(f)

    # Restructure into {wid: {HD1: path, HD2: path}}
    results = {}
    for (wid, hd_key), files in groups.items():
        best = _pick_best(files, dup_mode=dup_mode)
        if best is not None:
            results.setdefault(wid, {})[hd_key] = best
    return results


def _find_header_row(fpath):
    with open(fpath, "rb") as fh:
        for i, line in enumerate(fh):
            text = line.decode("utf-8-sig", errors="replace").strip()
            if text.startswith("Serial No"):
                return i
    return -1


def read_raw_table(fpath):
    header_idx = _find_header_row(fpath)
    if header_idx < 0:
        return pl.DataFrame()

    with open(fpath, "rb") as fh:
        lines = fh.readlines()

    header_cols = [
        str(c).strip() if c else f"_col{i}"
        for i, c in enumerate(lines[header_idx].decode("utf-8-sig", errors="replace").strip().split(","))
    ]

    data_rows = []
    for line in lines[header_idx + 1:]:
        text = line.decode("utf-8-sig", errors="replace").strip()
        if not text:
            continue
        first_val = text.split(",")[0].strip()
        try:
            int(first_val)
            data_rows.append(text)
        except ValueError:
            continue

    if not data_rows:
        return pl.DataFrame()

    csv_text = ",".join(header_cols) + "\n" + "\n".join(data_rows) + "\n"
    df = pl.read_csv(
        csv_text.encode("utf-8"),
        has_header=True,
        infer_schema_length=0,
        null_values=["1.#INF", "-1.#INF", "1.#IND", "nan", "NaN", ""],
    )
    for c in df.columns:
        if df.schema[c] == pl.String:
            df = df.with_columns(pl.col(c).cast(pl.Float64, strict=False).alias(c))

    rename = {}
    for c in df.columns:
        if c in COL_MAP:
            rename[c] = COL_MAP[c]
    if rename:
        df = df.rename(rename)

    cast_int = [c for c in df.columns if c in ("SHOT_ROW", "SHOT_COL", "WIS_ROW", "WIS_COL")]
    if cast_int:
        df = df.with_columns([pl.col(c).cast(pl.Int32, strict=False) for c in cast_int])

    all_data_cols = set(DATA_COLS + SERVO_COLS + PTC_COLS + TR_ELGA_COLS + TR_ELGB_COLS + TW_ELGA_COLS + TW_ELGB_COLS)
    numeric_cols = [c for c in df.columns if c in all_data_cols]
    if numeric_cols:
        df = df.with_columns([pl.col(c).cast(pl.Float32, strict=False) for c in numeric_cols])

    return df


def compute_global_coords(df):
    max_wis_r = df["WIS_ROW"].max()
    if max_wis_r is None:
        max_wis_r = 8
    max_wis_c = df["WIS_COL"].max()
    if max_wis_c is None:
        max_wis_c = 8
    wis_span = max(max_wis_r, max_wis_c)

    return df.with_columns([
        ((pl.col("SHOT_ROW") - 1) * wis_span + pl.col("WIS_ROW")).alias("G_ROW"),
        ((pl.col("SHOT_COL") - 1) * wis_span + pl.col("WIS_COL")).alias("G_COL"),
    ])


def cop_limits_from_cols(df, cols, center_ratio=0.8, ul_sig=3.0, ll_sig=3.0,
                         custom_3sigma=None):
    """Compute COP (Center Of Process) limits for the given columns.

    When *custom_3sigma* is provided (not None), the computed 3*sigma
    band is replaced by the user-specified half-width so that limits
    become ``mean ± custom_3sigma`` instead of ``mean ± ul_sig*sig``.
    This avoids over-killing normal dies when the data-driven 3*sigma
    is too narrow.
    """
    s = df.select(cols).unpivot(value_name="v").get_column("v").drop_nulls()
    if s.len() == 0:
        return 0.0, 1.0
    tail = (1.0 - center_ratio) / 2.0
    lo = s.quantile(tail)
    hi = s.quantile(1.0 - tail)
    center = s.filter((s >= lo) & (s <= hi))
    if center.len() == 0:
        return s.min(), s.max()
    mean = center.mean()
    sig = center.std()
    if sig is None or sig == 0:
        return mean, mean
    if custom_3sigma is not None:
        return mean - custom_3sigma, mean + custom_3sigma
    return mean - ll_sig * sig, mean + ul_sig * sig


def defect_map(df, keys, cols, lower, upper, code, name):
    checked = [
        pl.when(pl.col(c).is_between(lower, upper))
        .then(pl.col(c))
        .otherwise(None)
        .alias(c)
        for c in cols
    ]
    return (
        df.with_columns(checked)
        .with_columns(
            pl.when(pl.any_horizontal(pl.col(cols).is_null()))
            .then(True)
            .otherwise(False)
            .alias("_is_ng")
        )
        .group_by(keys)
        .agg(pl.col("_is_ng").any().alias("_has_ng"))
        .with_columns(
            pl.when(pl.col("_has_ng")).then(pl.lit(code)).otherwise(pl.lit("OK")).alias(name)
        )
        .select(keys + [name])
    )


def elg_defect_map(df, keys, elga_cols, elgb_cols, elga_lim, elgb_lim):
    elga_ll, elga_ul = elga_lim
    elgb_ll, elgb_ul = elgb_lim
    # Only the center 6 ELG channels are used for judgment.
    all_cols = elga_cols + elgb_cols
    center_cols = [c for c in ELG_CENTER_COLS if c in all_cols]
    return (
        df.with_columns(
            [pl.when(pl.col(c).is_between(elga_ll, elga_ul))
             .then(pl.col(c)).otherwise(None).alias(c)
             for c in elga_cols]
        )
        .with_columns(
            [pl.when(pl.col(c).is_between(elgb_ll, elgb_ul))
             .then(pl.col(c)).otherwise(None).alias(c)
             for c in elgb_cols]
        )
        .with_columns(
            pl.when(pl.sum_horizontal(pl.col(center_cols).is_null()) > 2)
            .then(True)
            .otherwise(False)
            .alias("_is_elg")
        )
        .group_by(keys)
        .agg(pl.col("_is_elg").any().alias("_has_elg"))
        .with_columns(
            pl.when(pl.col("_has_elg")).then(pl.lit("ELG")).otherwise(pl.lit("OK")).alias("ELG_MAP")
        )
        .select(keys + ["ELG_MAP"])
    )


def merge_defect_priority(dfs, keys, map_cols):
    base = dfs[0]
    for df in dfs[1:]:
        base = base.join(df, on=keys, how="left")
    first_ng = pl.coalesce(
        [pl.when(pl.col(c) != "OK").then(pl.col(c)) for c in map_cols]
        + [pl.lit("OK")]
    )
    return base.with_columns(first_ng.alias("FINL_MAP")).select(keys + map_cols + ["FINL_MAP"])


def inductance_defect_map(df, keys, lower, upper):
    return (
        df.with_columns(
            pl.when(pl.col("L").is_between(lower, upper))
            .then(False)
            .otherwise(True)
            .alias("_is_ind")
        )
        .group_by(keys)
        .agg(pl.col("_is_ind").any().alias("has_ind"))
        .with_columns(
            pl.when(pl.col("has_ind")).then(pl.lit("IND")).otherwise(pl.lit("OK")).alias("IND_MAP")
        )
        .select(keys + ["IND_MAP"])
    )


def _compute_yield(final_map):
    total = final_map.height
    ok_count = final_map.filter(pl.col("FINL_MAP") == "OK").height
    yield_rate = ok_count / total * 100 if total > 0 else 0
    defect_counts = {}
    for code in ["IND", "DAT", "SEV", "PTC", "ELG"]:
        defect_counts[code] = final_map.filter(pl.col("FINL_MAP") == code).height
    return {
        "total": total,
        "ok_count": ok_count,
        "yield_rate": yield_rate,
        "defect_counts": defect_counts,
    }


def _compute_hf_sig(df, data_cols):
    """Compute HF_Sig (moving-range standard deviation) per die.

    For each die, take the 32 Data channel values (HD01..HD32), compute
    the 31 moving-range values MR_i = |HD_{i+1} - HD_i|, then take the
    sample standard deviation of those MR values.

    Returns a DataFrame with INFO_COLS + 'HF_Sig', or None if fewer
    than 2 data columns are available. Dies with any null channel value
    get HF_Sig = null.
    """
    cols = [c for c in data_cols if c in df.columns]
    if len(cols) < 2:
        return None

    sub = df.select(INFO_COLS + cols).clone()
    mr_cols = []
    for i in range(len(cols) - 1):
        mr_col = f"_mr_{i:02d}"
        sub = sub.with_columns(
            (pl.col(cols[i + 1]) - pl.col(cols[i])).abs().alias(mr_col)
        )
        mr_cols.append(mr_col)

    n = len(mr_cols)
    # Only compute for rows where ALL data values are non-null
    all_valid = pl.all_horizontal([pl.col(c).is_not_null() for c in cols])

    sub = sub.with_columns(
        pl.sum_horizontal([pl.col(c) for c in mr_cols]).alias("_mr_sum")
    )
    sub = sub.with_columns(
        (pl.col("_mr_sum") / n).alias("_mr_mean")
    )
    ssd = pl.sum_horizontal(
        [(pl.col(c) - pl.col("_mr_mean")) ** 2 for c in mr_cols]
    )
    sub = sub.with_columns(
        pl.when(all_valid)
        .then((ssd / (n - 1)).sqrt())
        .otherwise(None)
        .alias("HF_Sig")
    )
    return sub.select(INFO_COLS + ["HF_Sig"])


def _gradient_color(t):
    """Map *t* in [0, 1] to a green→yellow→red hex color string."""
    t = max(0.0, min(1.0, t))
    if t < 0.5:
        s = t * 2
        r = int(0x4C + (0xFF - 0x4C) * s)
        g = int(0xAF + (0xEB - 0xAF) * s)
        b = int(0x50 + (0x3B - 0x50) * s)
    else:
        s = (t - 0.5) * 2
        r = 0xFF
        g = int(0xEB + (0x17 - 0xEB) * s)
        b = int(0x3B + (0x44 - 0x3B) * s)
    return f"{r:02X}{g:02X}{b:02X}"


def process_single_file(fpath, sensor_type, limits, lr_files=None, eps_format=None):
    raw_df = read_raw_table(str(fpath))
    if raw_df.height == 0:
        return None, "Empty data"

    wafer_id = extract_wafer_id(fpath)
    elga_cols = list(TR_ELGA_COLS if sensor_type == "TR" else TW_ELGA_COLS)
    elgb_cols = list(TR_ELGB_COLS if sensor_type == "TR" else TW_ELGB_COLS)
    available = set(raw_df.columns)
    elga_cols = [c for c in elga_cols if c in available]
    elgb_cols = [c for c in elgb_cols if c in available]
    data_cols = [c for c in DATA_COLS if c in available]
    servo_cols = [c for c in SERVO_COLS if c in available]
    ptc_cols = [c for c in PTC_COLS if c in available]

    data = raw_df.select(INFO_COLS + data_cols)
    servo = raw_df.select(INFO_COLS + servo_cols)
    elg = raw_df.select(INFO_COLS + elga_cols + elgb_cols)
    ptc = raw_df.select(INFO_COLS + ptc_cols)

    # For TW EPS-V3 format the EPS-V3 file has no HD{01..32} resistance
    # columns; those are read from HD1-LR/HD2-LR files instead.
    use_lr_data = (sensor_type == "TW" and eps_format == "EPS-V3")

    center_ratio = (limits or {}).get("center_ratio", 0.8)
    try:
        center_ratio = float(center_ratio)
    except (TypeError, ValueError):
        center_ratio = 0.8

    use_custom = bool(limits) and limits.get("use_custom")
    csig_data = (limits or {}).get("custom_3sigma_data")
    csig_l = (limits or {}).get("custom_3sigma_l")
    csig_servo = (limits or {}).get("custom_3sigma_servo")
    csig_elg = (limits or {}).get("custom_3sigma_elg")

    def _resolve_pair(ll_key, ul_key, df, cols, fallback=(0.0, 1.0),
                      custom_3sigma=None):
        """Per-item limit resolution: if BOTH custom LL and UL are
        provided, use them; otherwise fall back to COP 3Sigma computed
        from the wafer's own data. Only runs COP when the data actually
        contains the requested columns (guards EPS-V3 where HD columns
        come from LR files, not the EPS file); otherwise returns
        ``fallback`` (None for Data so the LR path computes it)."""
        if use_custom and ll_key in limits and ul_key in limits:
            return (limits[ll_key], limits[ul_key])
        present = [c for c in cols if c in df.columns]
        if present:
            return cop_limits_from_cols(df, present, center_ratio=center_ratio,
                                        custom_3sigma=custom_3sigma)
        return fallback

    data_lim = _resolve_pair("data_ll", "data_ul", data, DATA_COLS,
                             fallback=None, custom_3sigma=csig_data)
    servo_lim = _resolve_pair("servo_ll", "servo_ul", servo, SERVO_COLS,
                              custom_3sigma=csig_servo)
    elga_lim = _resolve_pair("elga_ll", "elga_ul", elg, elga_cols,
                             custom_3sigma=csig_elg)
    elgb_lim = _resolve_pair("elgb_ll", "elgb_ul", elg, elgb_cols,
                             custom_3sigma=csig_elg)
    if use_custom and "ptc_ll" in limits:
        ptc_ll = limits["ptc_ll"]
    else:
        ptc_ll = 1_500_000
    if use_custom and "l_ll" in limits and "l_ul" in limits:
        l_lim = (limits["l_ll"], limits["l_ul"])
    else:
        l_lim = None

    map_dfs = []
    map_cols = []

    l_lim_used = None
    ind_mode_used = "3sigma"
    cv_threshold_used = None
    data_lim_used = data_lim
    data_mode_used = "3sigma"
    data_cv_used = None
    data_hf_sig_used = None

    # Defect check toggles (default all enabled)
    checks = (limits or {}).get("defect_checks", {})
    ind_enabled = checks.get("IND", True)
    dat_enabled = checks.get("DAT", True)
    sev_enabled = checks.get("SEV", True)
    ptc_enabled = checks.get("PTC", True)
    elg_enabled = checks.get("ELG", True)

    # DATA judgment mode (3sigma / cv / hf_sigma / both). Applies to TW
    # Data; also harmless for TR (default 3sigma preserves original
    # behavior).
    data_mode = (limits or {}).get("data_mode", "3sigma")
    data_cv = (limits or {}).get("data_cv")
    data_hf_sig = (limits or {}).get("data_hf_sig")
    if data_mode not in ("3sigma", "cv", "hf_sigma", "both"):
        data_mode = "3sigma"

    # IND from LR files L column (TW only)
    if sensor_type == "TW" and lr_files and ind_enabled:
        ind_mode = (limits or {}).get("ind_mode", "3sigma")
        cv_threshold = (limits or {}).get("ind_cv")
        if ind_mode not in ("3sigma", "cv", "both"):
            ind_mode = "3sigma"
        l_map = _build_inductance_map(
            lr_files, wafer_id, l_lim,
            ind_mode=ind_mode, cv_threshold=cv_threshold,
            center_ratio=center_ratio, custom_3sigma=csig_l,
        )
        if l_map is not None:
            map_dfs.append(l_map)
            map_cols.append("IND_MAP")
            l_lim_used = l_map._l_lim if hasattr(l_map, "_l_lim") else l_lim
            if hasattr(l_map, "_ind_mode"):
                ind_mode_used = l_map._ind_mode
            if hasattr(l_map, "_cv_threshold"):
                cv_threshold_used = l_map._cv_threshold

    # DATA defect map. r_pivoted is always built (when use_lr_data) so
    # the Raw_Data sheet can include HD columns even if DAT is disabled.
    r_pivoted = None
    if use_lr_data:
        if lr_files is None:
            return None, "EPS-V3 format requires HD1-LR/HD2-LR files but none were found"
        r_map, r_data_lim, r_pivoted, r_cv_used, r_mode_used, r_hf_sig_used = _build_resistance_data_map(
            lr_files, wafer_id, data_lim,
            center_ratio=center_ratio, mode=data_mode, cv_threshold=data_cv,
            hf_sig_threshold=data_hf_sig, custom_3sigma=csig_data,
        )
        if r_pivoted is None:
            return None, "No resistance (R) data found in HD1-LR/HD2-LR files"
        if dat_enabled and r_map is not None:
            map_dfs.append(r_map)
            map_cols.append("DATA_MAP")
        if r_data_lim is not None:
            data_lim_used = r_data_lim
        data_mode_used = r_mode_used
        data_cv_used = r_cv_used
        data_hf_sig_used = r_hf_sig_used
    else:
        if dat_enabled and data_cols:
            d_map, d_lim_used, d_cv_used, d_hf_sig_used = _build_multi_mode_defect_map(
                data, INFO_COLS, DATA_COLS, data_lim,
                mode=data_mode, cv_threshold=data_cv,
                center_ratio=center_ratio, code="DAT", map_col="DATA_MAP",
                hf_sig_threshold=data_hf_sig, custom_3sigma=csig_data,
            )
            if d_map is not None:
                map_dfs.append(d_map)
                map_cols.append("DATA_MAP")
            if d_lim_used is not None:
                data_lim_used = d_lim_used
            data_mode_used = data_mode
            data_cv_used = d_cv_used
            data_hf_sig_used = d_hf_sig_used

    if sev_enabled and servo_cols:
        servo_map = defect_map(servo, INFO_COLS, SERVO_COLS, *servo_lim, "SEV", "SERVO_MAP")
        map_dfs.append(servo_map)
        map_cols.append("SERVO_MAP")

    if ptc_enabled and ptc_cols:
        ptc_map = defect_map(ptc, INFO_COLS, PTC_COLS, ptc_ll, float("inf"), "PTC", "PTC_MAP")
        map_dfs.append(ptc_map)
        map_cols.append("PTC_MAP")

    if elg_enabled and (elga_cols or elgb_cols):
        elg_map = elg_defect_map(elg, INFO_COLS, elga_cols, elgb_cols, elga_lim, elgb_lim)
        map_dfs.append(elg_map)
        map_cols.append("ELG_MAP")

    # If no defect checks are enabled, all dies are OK.
    if map_dfs:
        final_map = merge_defect_priority(map_dfs, keys=INFO_COLS, map_cols=map_cols)
    else:
        # Build a final_map with just INFO_COLS and all-OK FINL_MAP
        if use_lr_data and r_pivoted is not None:
            base = r_pivoted.select(INFO_COLS)
        elif data_cols:
            base = data.select(INFO_COLS)
        elif servo_cols:
            base = servo.select(INFO_COLS)
        else:
            base = raw_df.select(INFO_COLS)
        final_map = base.with_columns(pl.lit("OK").alias("FINL_MAP"))

    chipid_df = load_chipid_map(sensor_type)
    if chipid_df is not None:
        final_map = final_map.join(chipid_df, on=INFO_COLS, how="left")
        # Dies without a valid ChipID do not participate in yield
        # statistics or map display (edge/test dies not in the ChipID
        # mapping table).
        final_map = final_map.filter(pl.col("ChipID").is_not_null())
        chip_rows = [parse_chipid_rowcol(c)[0] for c in final_map["ChipID"].to_list()]
        chip_cols_list = [parse_chipid_rowcol(c)[1] for c in final_map["ChipID"].to_list()]
        final_map = final_map.with_columns([
            pl.Series("CHIP_ROW", chip_rows),
            pl.Series("CHIP_COL", chip_cols_list),
        ])

    # Build detailed data table: WaferID + ChipID + individual defect maps
    detail_df = final_map.select(
        (INFO_COLS + ["ChipID"] + map_cols + ["FINL_MAP"])
        if "ChipID" in final_map.columns else
        (INFO_COLS + map_cols + ["FINL_MAP"])
    )
    detail_df = detail_df.with_columns(pl.lit(wafer_id).alias("WaferID"))
    if "IND_MAP" in map_cols:
        detail_cols = ["WaferID", "ChipID", "IND_MAP", "DATA_MAP", "SERVO_MAP", "PTC_MAP", "ELG_MAP", "FINL_MAP"]
    else:
        detail_cols = ["WaferID", "ChipID", "DATA_MAP", "SERVO_MAP", "PTC_MAP", "ELG_MAP", "FINL_MAP"]
    available_detail = [c for c in detail_cols if c in detail_df.columns]
    detail_df = detail_df.select(available_detail)

    # Build raw data table for the Raw_Data sheet:
    # index = WaferID + ChipID; columns = L (TW only), Data (HD01..HD32),
    # Servo (S1, S2), ELG (RA*, RB*). PTC is intentionally excluded.
    raw_select = [c for c in (INFO_COLS + data_cols + servo_cols + elga_cols + elgb_cols)
                  if c in raw_df.columns]
    raw_data_df = raw_df.select(raw_select)

    # Data (HD01..HD32): for EPS-V3 these come from LR R column
    if use_lr_data and r_pivoted is not None:
        # Drop any HD cols that might have come from the EPS file (none
        # for EPS-V3) then join the LR-pivoted R wide table.
        existing_hd = [c for c in DATA_COLS if c in raw_data_df.columns]
        if existing_hd:
            raw_data_df = raw_data_df.drop(existing_hd)
        raw_data_df = raw_data_df.join(r_pivoted, on=INFO_COLS, how="left")

    # L (L01..L32) for TW, pivoted from LR files
    if sensor_type == "TW" and lr_files:
        l_pivoted = _pivot_lr_wide(lr_files, wafer_id, "L", L_COLS)
        if l_pivoted is not None:
            raw_data_df = raw_data_df.join(l_pivoted, on=INFO_COLS, how="left")

    # ChipID join
    if chipid_df is not None:
        raw_data_df = raw_data_df.join(chipid_df, on=INFO_COLS, how="left")
        # Keep only dies with a valid ChipID (consistent with final_map).
        raw_data_df = raw_data_df.filter(pl.col("ChipID").is_not_null())
    raw_data_df = raw_data_df.with_columns(pl.lit(wafer_id).alias("WaferID"))

    # Reorder: WaferID, ChipID first, then INFO_COLS, then L (TW),
    # Data, Servo, ELG in that order; any other columns appended.
    front = ["WaferID", "ChipID"]
    preferred = front + INFO_COLS + L_COLS + DATA_COLS + SERVO_COLS + elga_cols + elgb_cols
    rest = [c for c in raw_data_df.columns if c not in preferred]
    ordered = [c for c in preferred if c in raw_data_df.columns] + rest
    raw_data_df = raw_data_df.select(ordered)

    # Compute HF_Sig (moving-range standard deviation of HD01..HD32)
    # per die. ``raw_data_df`` carries the canonical HD01..HD32 columns
    # (from the EPS file for EPSH, or from the LR-pivoted R table for
    # EPS-V3), so it is the right input regardless of format.
    hf_sig_df = _compute_hf_sig(raw_data_df, DATA_COLS)
    if hf_sig_df is not None:
        # Attach ChipID / CHIP_ROW / CHIP_COL from final_map so the
        # Excel sheet can lay values out in wafer coordinates.
        pos_cols = [c for c in ["ChipID", "CHIP_ROW", "CHIP_COL"]
                    if c in final_map.columns]
        if pos_cols:
            hf_sig_df = hf_sig_df.join(
                final_map.select(INFO_COLS + pos_cols),
                on=INFO_COLS, how="left",
            )
        hf_sig_df = hf_sig_df.with_columns(pl.lit(wafer_id).alias("WaferID"))

    # Override Data_LL / Data_UL with the actual min/max of Data values
    # from dies that passed the judgment (OK), regardless of mode.
    ok_map_col = "DATA_MAP" if "DATA_MAP" in final_map.columns else "FINL_MAP"
    ok_dies = final_map.filter(pl.col(ok_map_col) == "OK").select(INFO_COLS)
    data_cols_present = [c for c in DATA_COLS if c in raw_data_df.columns]
    if data_cols_present and ok_dies.height > 0:
        ok_data = raw_data_df.join(ok_dies, on=INFO_COLS, how="inner")
        long = ok_data.unpivot(on=data_cols_present, index=INFO_COLS)
        vals = long["value"].drop_nulls()
        if vals.len() > 0:
            data_lim_used = (float(vals.min()), float(vals.max()))

    result_limits = {
        "file": fpath.name,
        "data_lim": data_lim_used,
        "servo_lim": servo_lim,
        "elga_lim": elga_lim,
        "elgb_lim": elgb_lim,
    }
    if l_lim_used:
        result_limits["l_lim"] = l_lim_used
    if ind_mode_used != "3sigma":
        result_limits["ind_mode"] = ind_mode_used
    if cv_threshold_used is not None:
        result_limits["ind_cv"] = cv_threshold_used
    if data_mode_used != "3sigma":
        result_limits["data_mode"] = data_mode_used
    if data_cv_used is not None:
        result_limits["data_cv"] = data_cv_used
    if data_hf_sig_used is not None:
        result_limits["data_hf_sig"] = data_hf_sig_used
    if csig_data is not None:
        result_limits["custom_3sigma_data"] = csig_data
    if csig_l is not None:
        result_limits["custom_3sigma_l"] = csig_l
    if csig_servo is not None:
        result_limits["custom_3sigma_servo"] = csig_servo
    if csig_elg is not None:
        result_limits["custom_3sigma_elg"] = csig_elg

    return {
        "final_map": final_map,
        "detail_df": detail_df,
        "raw_data_df": raw_data_df,
        "hf_sig_df": hf_sig_df,
        "limits": result_limits,
    }, None


def _build_multi_mode_defect_map(df, keys, cols, lim, mode="3sigma",
                                  cv_threshold=None, center_ratio=0.8,
                                  code="DAT", map_col="DATA_MAP",
                                  hf_sig_threshold=None, custom_3sigma=None):
    """Build a defect map supporting the same modes as IND plus an
    optional HF_Sigma mode for Data:

      - ``"3sigma"`` (default): WithinWafer 3Sigma. Compute COP 3-sigma
        limits across all values in ``cols`` on the wafer. Die is NG if
        any of its values is out of ``[lim_lower, lim_upper]``.
      - ``"cv"``: WithinChip CV%. For each die, compute
        CV% = std(cols) / mean(cols) * 100. Die is NG if CV% >
        ``cv_threshold``.
      - ``"hf_sigma"``: WithinChip moving-range standard deviation.
        For each die compute HF_Sig = std(|col_{i+1}-col_i|) across the
        ``len(cols)-1`` moving ranges; die is NG if HF_Sig >
        ``hf_sig_threshold``. Dies with any null channel get HF_Sig=null
        and are treated as OK (no judgment).
      - ``"both"``: Die is NG if either the 3-sigma OR the CV% check fails.

    Returns ``(map_df, lim_used, cv_used, hf_sig_used)`` where ``lim_used``
    is the [ll, ul] pair (always computed via COP for reference, even in
    cv/hf_sigma modes so the Limits sheet can show Data LL/UL), ``cv_used``
    is the CV% threshold applied (None otherwise), and ``hf_sig_used`` is
    the HF_Sig threshold applied (None otherwise).
    Returns ``(None, None, None, None)`` if ``df`` is empty / ``cols`` missing.
    """
    if df is None or df.height == 0 or not cols:
        return None, None, None, None

    # Always compute COP 3Sigma limits for reference (Limits sheet),
    # even when the 3Sigma map isn't used for judgment.
    if lim is None:
        lim = cop_limits_from_cols(df, cols, center_ratio=center_ratio,
                                   custom_3sigma=custom_3sigma)
    lim_used = lim

    # Only build the 3Sigma defect map when the mode uses it.
    map_3s = None
    if mode in ("3sigma", "both"):
        map_3s = defect_map(df, keys, cols, *lim, code, f"_{code}_3S")

    map_cv = None
    cv_used = None
    if mode in ("cv", "both") and cv_threshold is not None:
        long = df.unpivot(on=cols, index=keys, variable_name="_ch", value_name="_v")
        stats = long.group_by(keys).agg(
            pl.col("_v").mean().alias("_m"),
            pl.col("_v").std().alias("_s"),
        )
        stats = stats.with_columns(
            (pl.col("_s") / pl.col("_m") * 100.0).alias("_cv")
        )
        map_cv = (
            stats.with_columns(
                pl.when(pl.col("_cv").is_null())
                .then(pl.lit("OK"))
                .otherwise(
                    pl.when(pl.col("_cv") > cv_threshold)
                    .then(pl.lit(code))
                    .otherwise(pl.lit("OK"))
                )
                .alias(f"_{code}_CV")
            )
            .select(keys + [f"_{code}_CV"])
        )
        cv_used = cv_threshold

    map_hf = None
    hf_sig_used = None
    if mode == "hf_sigma" and hf_sig_threshold is not None:
        # Reuse the per-die HF_Sig computation; only dies where every
        # channel is non-null get a value, others fall back to OK.
        hf_df = _compute_hf_sig(df, cols)
        if hf_df is not None and "HF_Sig" in hf_df.columns:
            map_hf = (
                hf_df.with_columns(
                    pl.when(pl.col("HF_Sig").is_null())
                    .then(pl.lit("OK"))
                    .otherwise(
                        pl.when(pl.col("HF_Sig") > hf_sig_threshold)
                        .then(pl.lit(code))
                        .otherwise(pl.lit("OK"))
                    )
                    .alias(f"_{code}_HF")
                )
                .select(keys + [f"_{code}_HF"])
            )
            hf_sig_used = hf_sig_threshold

    if map_3s is not None and map_cv is not None:
        result = map_3s.join(map_cv, on=keys, how="left")
        result = result.with_columns(
            pl.when(
                (pl.col(f"_{code}_3S") == code) | (pl.col(f"_{code}_CV") == code)
            ).then(pl.lit(code)).otherwise(pl.lit("OK")).alias(map_col)
        ).select(keys + [map_col])
    elif map_3s is not None:
        result = map_3s.rename({f"_{code}_3S": map_col})
    elif map_cv is not None:
        result = map_cv.rename({f"_{code}_CV": map_col})
    elif map_hf is not None:
        result = map_hf.rename({f"_{code}_HF": map_col})
    else:
        result = df.select(keys).with_columns(pl.lit("OK").alias(map_col))

    return result, lim_used, cv_used, hf_sig_used


def _build_inductance_map(lr_files, wafer_id, l_lim, ind_mode="3sigma",
                          cv_threshold=None, center_ratio=0.8,
                          custom_3sigma=None):
    """Build IND defect map from L (inductance) values in HD1-LR/HD2-LR.

    Delegates to ``_build_multi_mode_defect_map`` with code="IND".
    IND does not use the HF_Sigma mode, so ``hf_sig_threshold`` is left
    at its default (None) and the returned ``hf_sig_used`` is ignored.
    """
    l_pivoted = _pivot_lr_wide(lr_files, wafer_id, "L", L_COLS)
    if l_pivoted is None:
        return None

    result, lim_used, cv_used, _hf_sig_used = _build_multi_mode_defect_map(
        l_pivoted, INFO_COLS, L_COLS, l_lim,
        mode=ind_mode, cv_threshold=cv_threshold,
        center_ratio=center_ratio, code="IND", map_col="IND_MAP",
        custom_3sigma=custom_3sigma,
    )
    if result is None:
        return None

    result._l_lim = lim_used if ind_mode in ("3sigma", "both") else None
    result._ind_mode = ind_mode
    result._cv_threshold = cv_used if ind_mode in ("cv", "both") else None
    return result


def _pivot_lr_wide(lr_files, wafer_id, value_col, out_cols):
    """Pivot HD1-LR / HD2-LR rows into a wide table: one row per die
    with channel values spread across ``out_cols`` columns.

    Channel mapping (same for R and L):
      HD1 even WIS_COL 2,4,...,32  -> out_cols[0..15]
      HD2 odd  WIS_COL 3,5,...,33  -> out_cols[16..31]

    ``value_col`` is the column name to read from the LR file (e.g.
    ``"R_L"`` for resistance, ``"L"`` for inductance).

    Returns a DataFrame with INFO_COLS + out_cols, or None if no data.
    """
    wafer_lr = lr_files.get(wafer_id, {})
    die_keys = ["SHOT_ROW", "SHOT_COL", "WIS_ROW", "WIS_COL"]

    col_by_wiscol = {}
    for i in range(1, 17):
        col_by_wiscol[2 * i] = out_cols[i - 1]          # HD1 even -> 0..15
    for i in range(1, 17):
        col_by_wiscol[2 * i + 1] = out_cols[i + 15]      # HD2 odd  -> 16..31

    all_rows = []
    for hd_key in ["HD1", "HD2"]:
        fpath = wafer_lr.get(hd_key)
        if fpath is None:
            continue
        try:
            df = read_raw_table(str(fpath))
            if df.height == 0 or value_col not in df.columns:
                continue
            sub = df.select(die_keys + [value_col])
            all_rows.append(sub)
        except Exception:
            continue

    if not all_rows:
        return None
    combined = pl.concat(all_rows, how="diagonal")
    if combined.height == 0:
        return None

    combined = combined.with_columns(
        pl.col("WIS_COL").replace_strict(col_by_wiscol, default=None).alias("_ch_name")
    )
    combined = combined.filter(
        pl.col("_ch_name").is_not_null() & pl.col(value_col).is_not_null()
    )
    pivoted = combined.pivot(
        on="_ch_name",
        index=["SHOT_ROW", "SHOT_COL", "WIS_ROW"],
        values=value_col,
        aggregate_function="first",
    )
    for c in out_cols:
        if c not in pivoted.columns:
            pivoted = pivoted.with_columns(pl.lit(None, dtype=pl.Float32).alias(c))
    pivoted = pivoted.with_columns(pl.lit(1).cast(pl.Int32).alias("WIS_COL"))
    return pivoted.select(INFO_COLS + out_cols)


def _build_resistance_data_map(lr_files, wafer_id, data_lim, center_ratio=0.8,
                                mode="3sigma", cv_threshold=None,
                                hf_sig_threshold=None, custom_3sigma=None):
    """Build a DATA defect map from the resistance (R) column of HD1-LR
    and HD2-LR files.

    Used for TW EPS-V3 format, where the EPS-V3 file does not contain
    HD{01..32} resistance columns but only PTC + Servo + ELG data. The
    resistance values are read from the ``R`` column of the HD1-LR and
    HD2-LR CSV files instead.

    Supports the same modes as ``_build_multi_mode_defect_map``
    (3sigma / cv / hf_sigma / both).

    Returns ``(data_map_df, data_lim_used, pivoted_wide_df, cv_used,
    mode_used, hf_sig_used)`` where ``pivoted_wide_df`` has
    INFO_COLS + DATA_COLS and is reused for the Raw_Data sheet. Returns
    ``(None, None, None, None, None, None)`` if no R data.
    """
    pivoted = _pivot_lr_wide(lr_files, wafer_id, "R_L", DATA_COLS)
    if pivoted is None:
        return None, None, None, None, None, None

    result, lim_used, cv_used, hf_sig_used = _build_multi_mode_defect_map(
        pivoted, INFO_COLS, DATA_COLS, data_lim,
        mode=mode, cv_threshold=cv_threshold,
        center_ratio=center_ratio, code="DAT", map_col="DATA_MAP",
        hf_sig_threshold=hf_sig_threshold, custom_3sigma=custom_3sigma,
    )
    return result, lim_used, pivoted, cv_used, mode, hf_sig_used


def process_lot(lot_no, sensor_type, limits=None, eps_format=None, dup_mode="size"):
    files = find_epsh_files(lot_no, sensor_type, eps_format=eps_format, dup_mode=dup_mode)
    if not files:
        return None, "No EPS data files found"

    lr_files = find_lr_files(lot_no, sensor_type, dup_mode=dup_mode)

    wafer_data = {}
    all_limits = []
    errors = []

    for fpath in files:
        wafer_id = extract_wafer_id(fpath)
        if wafer_id in wafer_data:
            continue
        try:
            result, err = process_single_file(fpath, sensor_type, limits, lr_files=lr_files, eps_format=eps_format)
            if err:
                errors.append(f"{fpath.name}: {err}")
                continue
            result["wafer_id"] = wafer_id
            result["file"] = str(fpath)
            yld = _compute_yield(result["final_map"])
            result.update(yld)
            wafer_data[wafer_id] = result
            all_limits.append(result["limits"])
        except Exception as e:
            errors.append(f"{fpath.name}: {e}")

    if not wafer_data:
        return None, "No valid data\n" + "\n".join(errors)

    total_ok = sum(w["ok_count"] for w in wafer_data.values())
    total_all = sum(w["total"] for w in wafer_data.values())
    total_yld = {
        "total": total_all,
        "ok_count": total_ok,
        "yield_rate": total_ok / total_all * 100 if total_all > 0 else 0,
        "defect_counts": {},
    }
    for code in ["IND", "DAT", "SEV", "PTC", "ELG"]:
        total_yld["defect_counts"][code] = sum(
            w["defect_counts"].get(code, 0) for w in wafer_data.values()
        )

    all_maps = [w["final_map"] for w in wafer_data.values()]
    combined = pl.concat(all_maps, how="diagonal")

    return {
        "wafer_data": wafer_data,
        "combined_map": combined,
        "total_yield": total_yld,
        "limits": all_limits,
        "sensor_type": sensor_type,
        "lot_no": lot_no,
        "eps_format": eps_format,
        "errors": errors,
    }, None


def generate_wafer_map_data(final_map):
    if "CHIP_ROW" not in final_map.columns or "CHIP_COL" not in final_map.columns:
        return None
    rows = final_map.get_column("CHIP_ROW").to_list()
    cols = final_map.get_column("CHIP_COL").to_list()
    maps_data = final_map.get_column("FINL_MAP").to_list()
    chipids = (final_map.get_column("ChipID").to_list()
               if "ChipID" in final_map.columns else [None] * len(rows))

    if not rows:
        return None

    r_min, r_max = min(rows), max(rows)
    c_min, c_max = min(cols), max(cols)

    grid = {}
    chipid_grid = {}
    for r, c, m, cid in zip(rows, cols, maps_data, chipids):
        grid[(r, c)] = m
        chipid_grid[(r, c)] = cid

    return grid, r_min, r_max, c_min, c_max, chipid_grid


def plot_wafer_map(ax, grid_data, title=""):
    if grid_data is None:
        ax.text(0.5, 0.5, "No Data", ha="center", va="center", fontsize=14)
        ax.set_title(title)
        return

    grid, r_min, r_max, c_min, c_max, _chipid_grid = grid_data
    n_rows = r_max - r_min + 1
    n_cols = c_max - c_min + 1

    for r in range(r_min, r_max + 1):
        for c in range(c_min, c_max + 1):
            val = grid.get((r, c), None)
            color = DEFECT_COLORS.get(val, "#E0E0E0") if val is not None else "#E0E0E0"
            i = r - r_min
            j = c - c_min
            rect = plt.Rectangle(
                (j, n_rows - 1 - i), 1, 1,
                facecolor=color, edgecolor="white", linewidth=0.3
            )
            ax.add_patch(rect)

    ax.set_xlim(0, n_cols)
    ax.set_ylim(0,n_rows)
    # ChipCol only has ~8 columns while ChipRow has ~176 rows, so an
    # "equal" data aspect makes the map a very tall narrow strip. Use
    # auto data aspect so the map fills the axes box, and force the
    # axes box itself to be square (box_aspect=1) so the DefectMap
    # displays as a square regardless of the data dimensions.
    ax.set_aspect("auto")
    ax.set_box_aspect(1)
    ax.set_xlabel("ChipCol")
    ax.set_ylabel("ChipRow")
    ax.set_title(title, fontsize=10)
    ax.tick_params(labelsize=6)


def add_legend(ax):
    handles = [
        mpatches.Patch(color=DEFECT_COLORS[k], label=k)
        for k in ["OK", "IND", "DAT", "SEV", "PTC", "ELG"]
    ]
    ax.legend(handles=handles, loc="upper right", fontsize=7, framealpha=0.8)


def write_excel(result, save_path, wafer_id=None):
    """Write an Excel workbook for the lot result.

    If ``wafer_id`` is None, all wafers in the lot are written into a
    single combined workbook (Summary / Limits / Map_{wid} / Detail Data).

    If ``wafer_id`` is given, only that wafer is written: the Summary
    sheet contains only that wafer's row (plus a TOTAL row), the Limits
    sheet only that wafer's limits, a single ``Map`` sheet, and the
    Detail sheet only that wafer's detail rows. This is used to produce
    one independent Excel file per WaferID.
    """
    wb = Workbook()

    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    ws = wb.active
    ws.title = "Yield Summary"

    lot_no = result["lot_no"]
    sensor_type = result["sensor_type"]
    wafer_data = result["wafer_data"]
    total_yld = result["total_yield"]

    if wafer_id is not None:
        wafer_data = {wid: w for wid, w in wafer_data.items() if wid == wafer_id}

    row = 1
    for c, v in enumerate(["Lot No", lot_no], 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = header_font if c == 1 else None
        cell.border = thin_border
        cell.alignment = center_align
    row += 1
    for c, v in enumerate(["Type", sensor_type], 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = header_font if c == 1 else None
        cell.border = thin_border
        cell.alignment = center_align
    row += 1
    row += 1

    headers = ["Wafer ID", "Total", "Pass", "Yield", "IND", "DAT", "SEV", "PTC", "ELG"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    row += 1

    for wid, w in wafer_data.items():
        dc = w["defect_counts"]
        vals = [
            wid, w["total"], w["ok_count"], f"{w['yield_rate']:.2f}%",
            dc.get("IND", 0), dc.get("DAT", 0), dc.get("SEV", 0),
            dc.get("PTC", 0), dc.get("ELG", 0),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.alignment = center_align
            cell.border = thin_border
        row += 1

    # TOTAL row: use only the selected wafers when filtered
    sel_total = sum(w["total"] for w in wafer_data.values())
    sel_ok = sum(w["ok_count"] for w in wafer_data.values())
    sel_defects = {}
    for code in ["IND", "DAT", "SEV", "PTC", "ELG"]:
        sel_defects[code] = sum(w["defect_counts"].get(code, 0) for w in wafer_data.values())
    sel_yield = sel_ok / sel_total * 100 if sel_total > 0 else 0

    total_vals = [
        "TOTAL", sel_total, sel_ok, f"{sel_yield:.2f}%",
        sel_defects.get("IND", 0), sel_defects.get("DAT", 0), sel_defects.get("SEV", 0),
        sel_defects.get("PTC", 0), sel_defects.get("ELG", 0),
    ]
    for c, v in enumerate(total_vals, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = Font(bold=True)
        cell.alignment = center_align
        cell.border = thin_border
    row += 1

    for i in range(1, 15):
        ws.column_dimensions[get_column_letter(i)].width = 14

    if result.get("limits"):
        ws_lim = wb.create_sheet("Limits")
        lim_headers = ["Wafer ID", "File", "L LL", "L UL",
                       "Data LL", "Data UL", "Servo LL", "Servo UL",
                       "ELGA LL", "ELGA UL", "ELGB LL", "ELGB UL",
                       "Data HF_Sig UL", "Data CV% UL",
                       "Data 3Sig", "L 3Sig", "Servo 3Sig", "ELG 3Sig"]
        for c, h in enumerate(lim_headers, 1):
            cell = ws_lim.cell(row=1, column=c, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        for r_idx, (wid, w) in enumerate(wafer_data.items(), 2):
            lim = w["limits"]
            vals = [
                wid, lim["file"],
                _fmt_lim(lim.get("l_lim"), 0),
                _fmt_lim(lim.get("l_lim"), 1),
                _fmt_lim(lim.get("data_lim"), 0),
                _fmt_lim(lim.get("data_lim"), 1),
                _fmt_lim(lim.get("servo_lim"), 0),
                _fmt_lim(lim.get("servo_lim"), 1),
                _fmt_lim(lim.get("elga_lim"), 0),
                _fmt_lim(lim.get("elga_lim"), 1),
                _fmt_lim(lim.get("elgb_lim"), 0),
                _fmt_lim(lim.get("elgb_lim"), 1),
                _fmt_scalar(lim.get("data_hf_sig")),
                _fmt_scalar(lim.get("data_cv")),
                _fmt_scalar(lim.get("custom_3sigma_data")),
                _fmt_scalar(lim.get("custom_3sigma_l")),
                _fmt_scalar(lim.get("custom_3sigma_servo")),
                _fmt_scalar(lim.get("custom_3sigma_elg")),
            ]
            for c, v in enumerate(vals, 1):
                cell = ws_lim.cell(row=r_idx, column=c, value=v)
                cell.alignment = center_align
                cell.border = thin_border

    defect_fills = {
        "OK": PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid"),
        "IND": PatternFill(start_color="FF1744", end_color="FF1744", fill_type="solid"),
        "DAT": PatternFill(start_color="F44336", end_color="F44336", fill_type="solid"),
        "SEV": PatternFill(start_color="FF9800", end_color="FF9800", fill_type="solid"),
        "PTC": PatternFill(start_color="9C27B0", end_color="9C27B0", fill_type="solid"),
        "ELG": PatternFill(start_color="2196F3", end_color="2196F3", fill_type="solid"),
    }

    for wid, w in wafer_data.items():
        sheet_name = "Map" if wafer_id is not None else f"Map_{wid}"
        ws_map = wb.create_sheet(sheet_name)
        final_map = w["final_map"]
        if "CHIP_ROW" not in final_map.columns:
            continue

        grid_data = generate_wafer_map_data(final_map)
        if grid_data is None:
            continue
        grid, r_min, r_max, c_min, c_max, chipid_grid = grid_data

        title_cell = ws_map.cell(row=1, column=1, value=f"{wid} - Yield: {w['yield_rate']:.2f}%")
        title_cell.font = Font(bold=True, size=12)

        for r in range(r_min, r_max + 1):
            for c in range(c_min, c_max + 1):
                val = grid.get((r, c), None)
                cid = chipid_grid.get((r, c), None)
                excel_row = r - r_min + 3
                excel_col = c - c_min + 2
                # Display "ChipID DefectCode" when both are available;
                # fall back to DefectCode only if ChipID is missing.
                if val:
                    display = f"{cid} {val}" if cid else val
                else:
                    display = ""
                cell = ws_map.cell(row=excel_row, column=excel_col, value=display)
                cell.alignment = center_align
                cell.border = thin_border
                if val and val in defect_fills:
                    cell.fill = defect_fills[val]
                    if val != "OK":
                        cell.font = Font(color="FFFFFF", bold=True, size=8)
                    else:
                        cell.font = Font(size=8)
                else:
                    cell.fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
                    cell.font = Font(size=8)

        for c in range(c_min, c_max + 1):
            cell = ws_map.cell(row=2, column=c - c_min + 2, value=str(c))
            cell.font = Font(bold=True, size=8)
            cell.alignment = center_align
        for r in range(r_min, r_max + 1):
            cell = ws_map.cell(row=r - r_min + 3, column=1, value=str(r))
            cell.font = Font(bold=True, size=8)
            cell.alignment = center_align

        ws_map.column_dimensions["A"].width = 6
        for c in range(c_min, c_max + 1):
            ws_map.column_dimensions[get_column_letter(c - c_min + 2)].width = 18
        # Row height tall enough for two-line content (ChipID + DefectCode)
        for r in range(r_min, r_max + 1):
            ws_map.row_dimensions[r - r_min + 3].height = 18
        ws_map.row_dimensions[2].height = 12

    # HF_Sig sheets: one per wafer, laid out by ChipRow/ChipCol with
    # cell color encoding the moving-range standard deviation value
    # (green=low → yellow → red=high, normalized within the wafer).
    for wid, w in wafer_data.items():
        hf_df = w.get("hf_sig_df")
        if hf_df is None or hf_df.height == 0:
            continue
        if "CHIP_ROW" not in hf_df.columns or "CHIP_COL" not in hf_df.columns:
            continue

        sheet_name = "HF_Sig" if wafer_id is not None else f"HF_Sig_{wid}"
        ws_sig = wb.create_sheet(sheet_name)

        rows = hf_df.get_column("CHIP_ROW").to_list()
        cols = hf_df.get_column("CHIP_COL").to_list()
        vals = hf_df.get_column("HF_Sig").to_list()
        if not rows:
            continue

        # Drop dies whose ChipID lookup failed (no CHIP_ROW/COL); they
        # cannot be positioned on the grid.
        triplets = [(r, c, v) for r, c, v in zip(rows, cols, vals)
                    if r is not None and c is not None]
        if not triplets:
            continue
        rows, cols, vals = zip(*triplets)
        rows, cols, vals = list(rows), list(cols), list(vals)

        r_min, r_max = min(rows), max(rows)
        c_min, c_max = min(cols), max(cols)

        # Normalize values within this wafer for coloring. Null entries
        # are skipped; if every value is null or all equal, default the
        # range to [0, 0] so the gradient falls back to the low color.
        finite = [v for v in vals if v is not None]
        if finite:
            v_min, v_max = min(finite), max(finite)
        else:
            v_min, v_max = 0.0, 0.0
        span = (v_max - v_min) or 1.0

        grid = {}
        for r, c, v in zip(rows, cols, vals):
            grid[(r, c)] = v

        title_cell = ws_sig.cell(row=1, column=1, value=f"{wid} - HF_Sig")
        title_cell.font = Font(bold=True, size=12)
        # Range hint beside the title for quick reference.
        if finite:
            rng_cell = ws_sig.cell(
                row=1, column=3,
                value=f"min={v_min:.4g}  max={v_max:.4g}",
            )
            rng_cell.font = Font(size=9, italic=True)

        for r in range(r_min, r_max + 1):
            for c in range(c_min, c_max + 1):
                val = grid.get((r, c), None)
                excel_row = r - r_min + 3
                excel_col = c - c_min + 2
                cell = ws_sig.cell(
                    row=excel_row, column=excel_col,
                    value=float(val) if val is not None else "",
                )
                cell.alignment = center_align
                cell.border = thin_border
                if val is not None:
                    t = (val - v_min) / span
                    color = _gradient_color(t)
                    cell.fill = PatternFill(
                        start_color=color, end_color=color, fill_type="solid"
                    )
                    # Pick a readable font color based on luminance.
                    lum = (int(color[0:2], 16) * 0.299 +
                           int(color[2:4], 16) * 0.587 +
                           int(color[4:6], 16) * 0.114)
                    cell.font = Font(
                        size=7,
                        color="FFFFFF" if lum < 140 else "000000",
                    )
                else:
                    cell.fill = PatternFill(
                        start_color="F0F0F0", end_color="F0F0F0",
                        fill_type="solid",
                    )
                    cell.font = Font(size=7)

        for c in range(c_min, c_max + 1):
            cell = ws_sig.cell(row=2, column=c - c_min + 2, value=str(c))
            cell.font = Font(bold=True, size=8)
            cell.alignment = center_align
        for r in range(r_min, r_max + 1):
            cell = ws_sig.cell(row=r - r_min + 3, column=1, value=str(r))
            cell.font = Font(bold=True, size=8)
            cell.alignment = center_align

        ws_sig.column_dimensions["A"].width = 6
        for c in range(c_min, c_max + 1):
            ws_sig.column_dimensions[get_column_letter(c - c_min + 2)].width = 10
        for r in range(r_min, r_max + 1):
            ws_sig.row_dimensions[r - r_min + 3].height = 10
        ws_sig.row_dimensions[2].height = 14

    ws_detail = wb.create_sheet("Detail Data")
    all_details = []
    for wid, w in wafer_data.items():
        if w.get("detail_df") is not None:
            all_details.append(w["detail_df"])
    if all_details:
        combined_detail = pl.concat(all_details, how="diagonal")
        detail_cols = combined_detail.columns
        for c_idx, h in enumerate(detail_cols, 1):
            cell = ws_detail.cell(row=1, column=c_idx, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        for r_idx in range(combined_detail.height):
            row_data = combined_detail.row(r_idx)
            for c_idx, val in enumerate(row_data, 1):
                cell = ws_detail.cell(row=r_idx + 2, column=c_idx, value=str(val) if val is not None else "")
                cell.alignment = center_align
                cell.border = thin_border
                col_name = detail_cols[c_idx - 1]
                if col_name in ("FINL_MAP", "IND_MAP", "DATA_MAP", "SERVO_MAP", "PTC_MAP", "ELG_MAP"):
                    if str(val) in defect_fills:
                        cell.fill = defect_fills[str(val)]

        for c_idx, col_name in enumerate(detail_cols, 1):
            ws_detail.column_dimensions[get_column_letter(c_idx)].width = 12

    # Raw_Data sheet: WaferID + ChipID index, L (TW) + Data + Servo + ELG
    # raw values (no PTC).
    ws_raw = wb.create_sheet("Raw_Data")
    all_raw = []
    for wid, w in wafer_data.items():
        if w.get("raw_data_df") is not None:
            all_raw.append(w["raw_data_df"])
    if all_raw:
        combined_raw = pl.concat(all_raw, how="diagonal")
        raw_cols = combined_raw.columns
        for c_idx, h in enumerate(raw_cols, 1):
            cell = ws_raw.cell(row=1, column=c_idx, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        for r_idx in range(combined_raw.height):
            row_data = combined_raw.row(r_idx)
            for c_idx, val in enumerate(row_data, 1):
                if val is None:
                    cell_val = ""
                elif isinstance(val, (int, float)):
                    cell_val = float(val)
                else:
                    cell_val = str(val)
                cell = ws_raw.cell(row=r_idx + 2, column=c_idx, value=cell_val)
                cell.alignment = center_align
                cell.border = thin_border

        for c_idx, col_name in enumerate(raw_cols, 1):
            ws_raw.column_dimensions[get_column_letter(c_idx)].width = 10

    wb.save(save_path)


def get_column_letter(col_idx):
    result = ""
    while col_idx > 0:
        col_idx -= 1
        result = chr(65 + col_idx % 26) + result
        col_idx //= 26
    return result


class YieldMapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Yield Map Generator")
        self.root.geometry("1280x860")
        self.root.minsize(960, 640)

        self.result = None
        self.config = load_config()
        self._wafer_var = tk.StringVar(value="ALL")
        self._cv_entry_enabled = True

        self._build_ui()
        self._load_saved_config()
        self.ind_mode_var.trace_add("write", lambda *_: self._update_ind_cv_state())
        self.data_mode_var.trace_add("write", lambda *_: (self._update_data_cv_state(), self._update_data_hf_sig_state()))

    def _build_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(main_frame, width=320)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        left_frame.pack_propagate(False)

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_input_panel(left_frame)
        self._build_config_panel(left_frame)
        self._build_result_panel(left_frame)
        self._build_chart_panel(right_frame)

    def _build_input_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Input", padding=8)
        frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(frame, text="Sensor Type:").pack(anchor=tk.W)
        self.type_var = tk.StringVar(value="TR")
        type_frame = ttk.Frame(frame)
        type_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(type_frame, text="TR", variable=self.type_var, value="TR",
                        command=self._on_type_change).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(type_frame, text="TW", variable=self.type_var, value="TW",
                        command=self._on_type_change).pack(side=tk.LEFT)

        # TW format toggle (EPSH old vs EPS-V3 new). Only relevant for TW.
        ttk.Label(frame, text="TW Format:").pack(anchor=tk.W, pady=(5, 0))
        self.twfmt_var = tk.StringVar(value="EPSH")
        twfmt_frame = ttk.Frame(frame)
        twfmt_frame.pack(fill=tk.X, pady=2)
        self.twfmt_radio_eps = ttk.Radiobutton(twfmt_frame, text="EPSH", variable=self.twfmt_var, value="EPSH")
        self.twfmt_radio_eps.pack(side=tk.LEFT, padx=(0, 10))
        self.twfmt_radio_v3 = ttk.Radiobutton(twfmt_frame, text="EPS-V3", variable=self.twfmt_var, value="EPS-V3")
        self.twfmt_radio_v3.pack(side=tk.LEFT)
        self._update_twfmt_state()

        ttk.Label(frame, text="Lot No:").pack(anchor=tk.W, pady=(5, 0))
        lot_frame = ttk.Frame(frame)
        lot_frame.pack(fill=tk.X, pady=2)
        self.lot_var = tk.StringVar()
        self.lot_combo = ttk.Combobox(lot_frame, textvariable=self.lot_var, width=15)
        self.lot_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(lot_frame, text="...", width=3, command=self._browse_lot).pack(side=tk.LEFT, padx=(2, 0))

        ttk.Label(frame, text="Wafer:").pack(anchor=tk.W, pady=(5, 0))
        self.wafer_combo = ttk.Combobox(frame, textvariable=self._wafer_var, width=20, state="readonly")
        self.wafer_combo.pack(fill=tk.X, pady=2)
        self.wafer_combo.bind("<<ComboboxSelected>>", self._on_wafer_change)

        # Defect checks: IND / DAT / SEV / PTC / ELG
        ttk.Label(frame, text="Defect Checks:").pack(anchor=tk.W, pady=(5, 0))
        checks_frame = ttk.Frame(frame)
        checks_frame.pack(fill=tk.X, pady=2)
        self.defect_vars = {}
        for code in ["IND", "DAT", "SEV", "PTC", "ELG"]:
            v = tk.BooleanVar(value=True)
            self.defect_vars[code] = v
            ttk.Checkbutton(checks_frame, text=code, variable=v).pack(side=tk.LEFT, padx=(0, 4))

        # Duplicate file selection mode: how to pick among repeated measurements
        ttk.Label(frame, text="Duplicate File:").pack(anchor=tk.W, pady=(5, 0))
        self.dup_mode_var = tk.StringVar(value="size")
        dup_frame = ttk.Frame(frame)
        dup_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(dup_frame, text="By Size", variable=self.dup_mode_var, value="size").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(dup_frame, text="By Newest", variable=self.dup_mode_var, value="newest").pack(side=tk.LEFT)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_frame, text="Generate", command=self._generate).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(btn_frame, text="Export Excel", command=self._export_excel).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

    def _build_config_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Limits Configuration", padding=8)
        frame.pack(fill=tk.X, pady=(0, 5))

        # IND mode (TW L inductance judgment)
        ttk.Label(frame, text="IND Mode (TW L):").pack(anchor=tk.W)
        self.ind_mode_var = tk.StringVar(value="3sigma")
        ind_frame = ttk.Frame(frame)
        ind_frame.pack(fill=tk.X, pady=2)
        self.ind_rb_3s = ttk.Radiobutton(ind_frame, text="3Sigma", variable=self.ind_mode_var, value="3sigma")
        self.ind_rb_3s.pack(side=tk.LEFT, padx=(0, 6))
        self.ind_rb_cv = ttk.Radiobutton(ind_frame, text="WIR_CV%", variable=self.ind_mode_var, value="cv")
        self.ind_rb_cv.pack(side=tk.LEFT, padx=(0, 6))
        self.ind_rb_both = ttk.Radiobutton(ind_frame, text="3Sigma+CV%", variable=self.ind_mode_var, value="both")
        self.ind_rb_both.pack(side=tk.LEFT)

        cv_row = ttk.Frame(frame)
        cv_row.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(cv_row, text="CV% UL", width=10).pack(side=tk.LEFT)
        self.ind_cv_entry = ttk.Entry(cv_row, width=12)
        self.ind_cv_entry.insert(0, "3.0")
        self.ind_cv_entry.pack(side=tk.LEFT, padx=(2, 0))
        self._update_ind_cv_state()

        # Data mode (Data resistance judgment). Same modes as IND plus
        # an additional HF_Sigma mode that judges NG based on the
        # moving-range standard deviation HF_Sig of the 32 channels.
        ttk.Label(frame, text="Data Mode:").pack(anchor=tk.W)
        self.data_mode_var = tk.StringVar(value="3sigma")
        data_mode_frame = ttk.Frame(frame)
        data_mode_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(data_mode_frame, text="3Sig", variable=self.data_mode_var, value="3sigma").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(data_mode_frame, text="WIR_CV%", variable=self.data_mode_var, value="cv").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(data_mode_frame, text="HF_Sig", variable=self.data_mode_var, value="hf_sigma").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Radiobutton(data_mode_frame, text="3Sig+CV%", variable=self.data_mode_var, value="both").pack(side=tk.LEFT)

        data_cv_row = ttk.Frame(frame)
        data_cv_row.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(data_cv_row, text="CV% UL", width=10).pack(side=tk.LEFT)
        self.data_cv_entry = ttk.Entry(data_cv_row, width=12)
        self.data_cv_entry.insert(0, "6.0")
        self.data_cv_entry.pack(side=tk.LEFT, padx=(2, 0))
        self._update_data_cv_state()

        # HF_Sig upper limit; only meaningful in HF_Sigma mode (and
        # reserved for future 'both' extension). Disabled otherwise.
        data_hf_row = ttk.Frame(frame)
        data_hf_row.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(data_hf_row, text="HF_Sig UL", width=10).pack(side=tk.LEFT)
        self.data_hf_sig_entry = ttk.Entry(data_hf_row, width=12)
        self.data_hf_sig_entry.insert(0, "0.5")
        self.data_hf_sig_entry.pack(side=tk.LEFT, padx=(2, 0))
        self._update_data_hf_sig_state()

        self.custom_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Use Custom Limits", variable=self.custom_var,
                        command=self._toggle_custom).pack(anchor=tk.W)

        cr_row = ttk.Frame(frame)
        cr_row.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(cr_row, text="Center Ratio", width=14).pack(side=tk.LEFT)
        self.center_ratio_var = tk.StringVar(value="0.8")
        cr_entry = ttk.Entry(cr_row, textvariable=self.center_ratio_var, width=8)
        cr_entry.pack(side=tk.LEFT, padx=(2, 0))

        # Custom 3Sigma per item type: blank entry = use COP 3Sigma.
        # Each value is the half-width (mean ± value) replacing 3*sigma.
        ttk.Label(frame, text="Custom 3Sig (blank=COP):").pack(anchor=tk.W, pady=(4, 0))
        csig_frame = ttk.Frame(frame)
        csig_frame.pack(fill=tk.X, pady=2)
        self.custom_3sig_entries = {}
        csig_items = [
            ("data_3sig", "Data"), ("l_3sig", "L"),
            ("servo_3sig", "Servo"), ("elg_3sig", "ELG"),
        ]
        for idx, (key, label) in enumerate(csig_items):
            r, c = divmod(idx, 2)
            ttk.Label(csig_frame, text=label, width=8).grid(
                row=r, column=c * 2, sticky=tk.W, padx=(0, 2), pady=1)
            entry = ttk.Entry(csig_frame, width=10)
            entry.grid(row=r, column=c * 2 + 1, padx=(0, 8), pady=1)
            self.custom_3sig_entries[key] = entry

        self.lim_frame = ttk.Frame(frame)
        self.lim_frame.pack(fill=tk.X, pady=(4, 0))

        self.lim_entries = {}
        lim_items = [
            ("l_ll", "L LL"), ("l_ul", "L UL"),
            ("data_ll", "Data LL"), ("data_ul", "Data UL"),
            ("servo_ll", "Servo LL"), ("servo_ul", "Servo UL"),
            ("elga_ll", "ELGA LL"), ("elga_ul", "ELGA UL"),
            ("elgb_ll", "ELGB LL"), ("elgb_ul", "ELGB UL"),
            ("ptc_ll", "PTC LL"),
        ]
        # Two-column grid layout so all custom limit entries fit
        # vertically and leave room for the Result panel below.
        for idx, (key, label) in enumerate(lim_items):
            r, c = divmod(idx, 2)
            ttk.Label(self.lim_frame, text=label, width=10).grid(
                row=r, column=c * 2, sticky=tk.W, padx=(0, 2), pady=1)
            entry = ttk.Entry(self.lim_frame, width=10)
            entry.grid(row=r, column=c * 2 + 1, padx=(0, 8), pady=1)
            self.lim_entries[key] = entry

        self._toggle_custom()

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_row, text="Save Config", command=self._save_config).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(btn_row, text="Load Config", command=self._load_config).pack(side=tk.LEFT)

    def _update_ind_cv_state(self):
        """Enable the CV% entry only when the IND mode uses CV
        ('cv' or 'both'). For '3sigma' it is disabled."""
        mode = self.ind_mode_var.get()
        state = "normal" if mode in ("cv", "both") else "disabled"
        self.ind_cv_entry.configure(state=state)

    def _update_data_cv_state(self):
        """Enable the Data CV% entry only when the Data mode uses CV
        ('cv' or 'both'). For '3sigma' / 'hf_sigma' it is disabled."""
        mode = self.data_mode_var.get()
        state = "normal" if mode in ("cv", "both") else "disabled"
        self.data_cv_entry.configure(state=state)

    def _update_data_hf_sig_state(self):
        """Enable the HF_Sig UL entry only when the Data mode is
        'hf_sigma'. Disabled otherwise."""
        mode = self.data_mode_var.get()
        state = "normal" if mode == "hf_sigma" else "disabled"
        self.data_hf_sig_entry.configure(state=state)

    def _build_result_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Result", padding=8)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        self.result_text = tk.Text(frame, height=14, width=38, font=("Consolas", 9), state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_chart_panel(self, parent):
        self.fig, self.axes = plt.subplots(1, 1, figsize=(8, 8))
        self.fig.subplots_adjust(left=0.06, right=0.94, top=0.94, bottom=0.05)

        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas.draw()

        toolbar_frame = ttk.Frame(canvas_frame)
        toolbar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        NavigationToolbar2Tk(self.canvas, toolbar_frame)

        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _on_type_change(self):
        self._update_twfmt_state()
        self._refresh_lots()

    def _update_twfmt_state(self):
        """Enable the TW format toggle only when TW is selected; for TR
        the toggle is disabled (TR always uses EPSH)."""
        is_tw = self.type_var.get() == "TW"
        state = "normal" if is_tw else "disabled"
        self.twfmt_radio_eps.configure(state=state)
        self.twfmt_radio_v3.configure(state=state)

    def _get_eps_format(self):
        """Return the EPS format to use. For TR, always None (auto EPSH).
        For TW, return the selected toggle value ('EPSH' or 'EPS-V3')."""
        if self.type_var.get() == "TR":
            return None
        return self.twfmt_var.get()

    def _refresh_lots(self):
        stype = self.type_var.get()
        lots = list_lots(stype)
        self.lot_combo["values"] = lots
        if lots:
            self.lot_combo.set(lots[-1])

    def _browse_lot(self):
        stype = self.type_var.get()
        root_dir = TR_ROOT if stype == "TR" else TW_ROOT
        lot_dir = filedialog.askdirectory(initialdir=root_dir, title="Select Lot Directory")
        if lot_dir:
            self.lot_var.set(Path(lot_dir).name)

    def _on_wafer_change(self, event=None):
        if not self.result:
            return
        self._display_chart()

    def _toggle_custom(self):
        state = "normal" if self.custom_var.get() else "disabled"
        for entry in self.lim_entries.values():
            entry.configure(state=state)

    def _get_limits(self):
        limits = {}
        # Center ratio always passed (applies to COP auto-calculation)
        cr_str = self.center_ratio_var.get().strip()
        try:
            limits["center_ratio"] = float(cr_str) if cr_str else 0.7
        except ValueError:
            messagebox.showwarning("Warning", f"Invalid Center Ratio: {cr_str}")
            return None
        # Defect check toggles (all on by default)
        limits["defect_checks"] = {code: v.get() for code, v in self.defect_vars.items()}
        if self.custom_var.get():
            limits["use_custom"] = True
            # Only record explicitly provided values; entries left blank
            # are omitted so process_single_file falls back to 3Sigma COP
            # for that specific item (per-LL/UL granularity).
            for key, entry in self.lim_entries.items():
                val = entry.get().strip()
                if not val:
                    continue
                try:
                    limits[key] = float(val)
                except ValueError:
                    messagebox.showwarning("Warning", f"Invalid value for {key}: {val}")
                    return None
        # IND mode + CV threshold are always passed (even without custom
        # limits), because the mode selection is independent of the
        # numeric limit values.
        if self.type_var.get() == "TW":
            limits["ind_mode"] = self.ind_mode_var.get()
            if limits["ind_mode"] in ("cv", "both"):
                cv_str = self.ind_cv_entry.get().strip()
                if not cv_str:
                    messagebox.showwarning("Warning", "Please enter CV% upper limit.")
                    return None
                try:
                    limits["ind_cv"] = float(cv_str)
                except ValueError:
                    messagebox.showwarning("Warning", f"Invalid CV% value: {cv_str}")
                    return None
        # Data mode + CV threshold (always passed, same as IND mode)
        limits["data_mode"] = self.data_mode_var.get()
        if limits["data_mode"] in ("cv", "both"):
            dcv_str = self.data_cv_entry.get().strip()
            if not dcv_str:
                messagebox.showwarning("Warning", "Please enter Data CV% upper limit.")
                return None
            try:
                limits["data_cv"] = float(dcv_str)
            except ValueError:
                messagebox.showwarning("Warning", f"Invalid Data CV% value: {dcv_str}")
                return None
        # HF_Sigma mode + HF_Sig UL (always passed; only enforced when
        # data_mode == 'hf_sigma').
        if limits["data_mode"] == "hf_sigma":
            dhf_str = self.data_hf_sig_entry.get().strip()
            if not dhf_str:
                messagebox.showwarning("Warning", "Please enter Data HF_Sig upper limit.")
                return None
            try:
                limits["data_hf_sig"] = float(dhf_str)
            except ValueError:
                messagebox.showwarning("Warning", f"Invalid Data HF_Sig value: {dhf_str}")
                return None
        # Custom 3Sigma per item type: blank = use COP 3Sigma.
        # Keys: custom_3sigma_data / _l / _servo / _elg.
        csig_keys = {
            "data_3sig": "custom_3sigma_data",
            "l_3sig": "custom_3sigma_l",
            "servo_3sig": "custom_3sigma_servo",
            "elg_3sig": "custom_3sigma_elg",
        }
        for gui_key, lim_key in csig_keys.items():
            val = self.custom_3sig_entries[gui_key].get().strip()
            if not val:
                continue
            try:
                limits[lim_key] = float(val)
            except ValueError:
                messagebox.showwarning("Warning", f"Invalid Custom 3Sig value: {val}")
                return None
        return limits

    def _save_config(self):
        cfg = {"type": self.type_var.get(), "custom": self.custom_var.get()}
        if self.type_var.get() == "TW":
            cfg["tw_format"] = self.twfmt_var.get()
        cfg["ind_mode"] = self.ind_mode_var.get()
        cfg["ind_cv"] = self.ind_cv_entry.get().strip()
        cfg["data_mode"] = self.data_mode_var.get()
        cfg["data_cv"] = self.data_cv_entry.get().strip()
        cfg["data_hf_sig"] = self.data_hf_sig_entry.get().strip()
        cfg["center_ratio"] = self.center_ratio_var.get()
        cfg["custom_3sig"] = {k: e.get().strip() for k, e in self.custom_3sig_entries.items()}
        cfg["defect_checks"] = {code: v.get() for code, v in self.defect_vars.items()}
        cfg["dup_mode"] = self.dup_mode_var.get()
        for key, entry in self.lim_entries.items():
            cfg[key] = entry.get().strip()
        save_config(cfg)
        messagebox.showinfo("Info", "Configuration saved.")

    def _load_config(self):
        cfg = load_config()
        if not cfg:
            messagebox.showinfo("Info", "No saved configuration found.")
            return
        self._apply_config(cfg)

    def _load_saved_config(self):
        cfg = load_config()
        if cfg:
            self._apply_config(cfg)

    def _apply_config(self, cfg):
        if "type" in cfg:
            self.type_var.set(cfg["type"])
        if "tw_format" in cfg:
            self.twfmt_var.set(cfg["tw_format"])
        if "ind_mode" in cfg:
            self.ind_mode_var.set(cfg["ind_mode"])
        if "ind_cv" in cfg:
            self.ind_cv_entry.delete(0, tk.END)
            self.ind_cv_entry.insert(0, cfg["ind_cv"])
        if "data_mode" in cfg:
            self.data_mode_var.set(cfg["data_mode"])
        if "data_cv" in cfg:
            self.data_cv_entry.delete(0, tk.END)
            self.data_cv_entry.insert(0, cfg["data_cv"])
        if "data_hf_sig" in cfg:
            self.data_hf_sig_entry.delete(0, tk.END)
            self.data_hf_sig_entry.insert(0, cfg["data_hf_sig"])
        if "center_ratio" in cfg:
            self.center_ratio_var.set(cfg["center_ratio"])
        if "custom_3sig" in cfg and isinstance(cfg["custom_3sig"], dict):
            for k, v in cfg["custom_3sig"].items():
                if k in self.custom_3sig_entries:
                    self.custom_3sig_entries[k].delete(0, tk.END)
                    self.custom_3sig_entries[k].insert(0, v)
        if "defect_checks" in cfg:
            for code, val in cfg["defect_checks"].items():
                if code in self.defect_vars:
                    self.defect_vars[code].set(val)
        if "custom" in cfg:
            self.custom_var.set(cfg["custom"])
        if "dup_mode" in cfg:
            self.dup_mode_var.set(cfg["dup_mode"])
        for key, entry in self.lim_entries.items():
            if key in cfg:
                entry.delete(0, tk.END)
                entry.insert(0, cfg[key])
        self._update_twfmt_state()
        self._update_ind_cv_state()
        self._update_data_cv_state()
        self._update_data_hf_sig_state()
        self._toggle_custom()

    def _generate(self):
        lot_no = self.lot_var.get().strip()
        if not lot_no:
            messagebox.showwarning("Warning", "Please enter a Lot No.")
            return

        stype = self.type_var.get()
        eps_format = self._get_eps_format()
        limits = self._get_limits()
        dup_mode = self.dup_mode_var.get()

        self.root.configure(cursor="watch")
        self.root.update()

        try:
            result, error = process_lot(lot_no, stype, limits, eps_format=eps_format, dup_mode=dup_mode)
            if error:
                messagebox.showerror("Error", error)
                return

            self.result = result

            wafer_ids = list(result["wafer_data"].keys())
            self.wafer_combo["values"] = ["ALL"] + wafer_ids
            if wafer_ids:
                self._wafer_var.set("ALL")

            self._display_result()
            self._display_chart()

            if not self.custom_var.get() and result.get("limits"):
                lim = result["limits"][-1]
                auto_vals = {
                    "data_ll": _fmt_lim(lim.get("data_lim"), 0),
                    "data_ul": _fmt_lim(lim.get("data_lim"), 1),
                    "servo_ll": _fmt_lim(lim.get("servo_lim"), 0),
                    "servo_ul": _fmt_lim(lim.get("servo_lim"), 1),
                    "elga_ll": _fmt_lim(lim.get("elga_lim"), 0),
                    "elga_ul": _fmt_lim(lim.get("elga_lim"), 1),
                    "elgb_ll": _fmt_lim(lim.get("elgb_lim"), 0),
                    "elgb_ul": _fmt_lim(lim.get("elgb_lim"), 1),
                }
                if lim.get("l_lim") is not None:
                    auto_vals["l_ll"] = _fmt_lim(lim.get("l_lim"), 0)
                    auto_vals["l_ul"] = _fmt_lim(lim.get("l_lim"), 1)
                for key, val in auto_vals.items():
                    self.lim_entries[key].delete(0, tk.END)
                    self.lim_entries[key].insert(0, val)
                    self.lim_entries[key].configure(state="disabled")
            elif self.custom_var.get() and result.get("limits"):
                # Custom mode: only fill entries the user left blank with
                # the computed 3Sigma values, so they can see what was used.
                # Entries the user already filled are preserved.
                lim = result["limits"][-1]
                computed_vals = {
                    "data_ll": _fmt_lim(lim.get("data_lim"), 0),
                    "data_ul": _fmt_lim(lim.get("data_lim"), 1),
                    "servo_ll": _fmt_lim(lim.get("servo_lim"), 0),
                    "servo_ul": _fmt_lim(lim.get("servo_lim"), 1),
                    "elga_ll": _fmt_lim(lim.get("elga_lim"), 0),
                    "elga_ul": _fmt_lim(lim.get("elga_lim"), 1),
                    "elgb_ll": _fmt_lim(lim.get("elgb_lim"), 0),
                    "elgb_ul": _fmt_lim(lim.get("elgb_lim"), 1),
                }
                if lim.get("l_lim") is not None:
                    computed_vals["l_ll"] = _fmt_lim(lim.get("l_lim"), 0)
                    computed_vals["l_ul"] = _fmt_lim(lim.get("l_lim"), 1)
                for key, val in computed_vals.items():
                    if not self.lim_entries[key].get().strip():
                        self.lim_entries[key].delete(0, tk.END)
                        self.lim_entries[key].insert(0, val)

        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.root.configure(cursor="")

    def _display_result(self):
        if not self.result:
            return
        r = self.result
        ty = r["total_yield"]
        lines = [
            f"Lot: {r['lot_no']}  Type: {r['sensor_type']}",
            f"{'='*36}",
            f"{'WaferID':<12}{'Total':>6}{'Pass':>6}{'Yield':>8}",
            f"{'-'*36}",
        ]
        for wid, w in r["wafer_data"].items():
            lines.append(f"{wid:<12}{w['total']:>6}{w['ok_count']:>6}{w['yield_rate']:>7.2f}%")
        lines.append(f"{'-'*36}")
        lines.append(f"{'TOTAL':<12}{ty['total']:>6}{ty['ok_count']:>6}{ty['yield_rate']:>7.2f}%")
        lines.append(f"{'='*36}")
        lines.append(f"Defect Breakdown (Total):")
        for code, cnt in ty["defect_counts"].items():
            rate = cnt / ty["total"] * 100 if ty["total"] > 0 else 0
            lines.append(f"  {code:4s}: {cnt:5d} ({rate:.2f}%)")
        if r.get("errors"):
            lines.append(f"\nWarnings:")
            for e in r["errors"][:5]:
                lines.append(f"  {e[:50]}")

        self.result_text.configure(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, "\n".join(lines))
        self.result_text.configure(state=tk.DISABLED)

    def _display_chart(self):
        if not self.result:
            return

        self.fig.clear()

        wafer_data = self.result["wafer_data"]
        selected = self._wafer_var.get()

        if selected == "ALL":
            n_wafers = len(wafer_data)
            if n_wafers <= 4:
                ncols = 2
            else:
                ncols = 3
            nrows = (n_wafers + ncols - 1) // ncols
            self.axes = self.fig.subplots(nrows, ncols, squeeze=False)
            for ax_row in self.axes:
                for ax in ax_row:
                    ax.set_visible(False)

            for idx, (wid, w) in enumerate(wafer_data.items()):
                r, c = divmod(idx, ncols)
                ax = self.axes[r][c]
                ax.set_visible(True)
                ax.clear()
                grid_data = generate_wafer_map_data(w["final_map"])
                title = f"{wid} Yield:{w['yield_rate']:.1f}% ({w['ok_count']}/{w['total']})"
                plot_wafer_map(ax, grid_data, title)
                add_legend(ax)
        else:
            self.axes = self.fig.subplots(1, 1)
            ax = self.axes
            ax.clear()
            w = wafer_data.get(selected)
            if w:
                grid_data = generate_wafer_map_data(w["final_map"])
                title = f"{self.result['lot_no']} {selected} Yield:{w['yield_rate']:.1f}% ({w['ok_count']}/{w['total']})"
                plot_wafer_map(ax, grid_data, title)
                add_legend(ax)

        self.fig.tight_layout()
        self.canvas.draw()

    def _export_excel(self):
        if not self.result:
            messagebox.showwarning("Warning", "Please generate result first.")
            return

        lot_no = self.result["lot_no"]
        stype = self.result["sensor_type"]
        wafer_data = self.result["wafer_data"]
        selected = self._wafer_var.get()

        # If a specific wafer is selected, export that single wafer to
        # one user-named Excel file. If "ALL" is selected, export each
        # wafer to its own independent Excel file in a chosen folder.
        if selected != "ALL":
            if selected not in wafer_data:
                messagebox.showwarning("Warning", f"Wafer {selected} not found.")
                return
            default_name = f"{lot_no}_{stype}_{selected}_YieldMap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_name,
            )
            if not save_path:
                return
            try:
                self.root.configure(cursor="watch")
                self.root.update()
                write_excel(self.result, save_path, wafer_id=selected)
                messagebox.showinfo("Success", f"Excel saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                self.root.configure(cursor="")
            return

        # ALL -> one independent Excel per WaferID in a chosen folder
        out_dir = filedialog.askdirectory(title="Select output folder for per-wafer Excel files")
        if not out_dir:
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            self.root.configure(cursor="watch")
            self.root.update()
            saved = []
            for wid in wafer_data.keys():
                fname = f"{lot_no}_{stype}_{wid}_YieldMap_{stamp}.xlsx"
                save_path = str(Path(out_dir) / fname)
                write_excel(self.result, save_path, wafer_id=wid)
                saved.append(save_path)
            messagebox.showinfo(
                "Success",
                f"Saved {len(saved)} Excel files (one per WaferID) to:\n{out_dir}",
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.root.configure(cursor="")


def main():
    root = tk.Tk()
    app = YieldMapApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
