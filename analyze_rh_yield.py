# -*- coding: utf-8 -*-
"""
RH 数据良率统计脚本 v2
- 读取目录下所有 xlsx 文件的 'RH' sheet
- LotNo/WaferID 从路径正则匹配, ChipID 从 RH sheet 的 Part ID 列读取
- 仅统计 ChipID 中含 'R' 的晶圆
- 对 34 个通道 (S1, H1~H32, S2) 计算 ±100Gs 和 ±200Gs 区间下 BJ%
  BJ% = Max(ABS(相邻点电阻差值)) / (区间内电阻极值)  (Forward / Reverse 各算一次，取最大)
- 输出: 原始数据 CSV (LotNo, WaferID, ChipID, ChannelID, BJ%_100Gs, BJ%_200Gs)
        交互式 HTML (LotNo/WaferID 筛选, BJ%分布对比, 可调 UL 通道/Chip 良率)
"""
import os
import re
import glob
import json
import numpy as np
import pandas as pd
from openpyxl import load_workbook

ROOT_DIR = r'\\server\share\Raw data'  # 数据源目录, 按实际环境修改
OUT_DIR = os.path.join(os.path.expanduser('~'), 'RH_Yield_Output')
UL = 5.0

CHANNEL_ORDER = ['S1'] + [f'H{i}' for i in range(1, 33)] + ['S2']


# ── 路径解析 ──────────────────────────────────────────────

def parse_lot_wafer(filepath):
    """从父目录名正则匹配 LotNo 和 WaferID"""
    parts = filepath.replace('/', '\\').split('\\')
    parent = parts[-2]

    # Pattern 1: "WaferID (LotNo)"  e.g. MD24J0034 (259079)
    m = re.match(r'^(\S+)\s+\((\S+)\)$', parent)
    if m:
        return m.group(2), m.group(1)

    # Pattern 2: "LotNo WaferID"  e.g. 263021 MD25A1518
    m = re.match(r'^(\S+)\s+(.+)$', parent)
    if m:
        return m.group(1), m.group(2)

    # Pattern 3: "LotNo_WaferID" where WaferID has no trailing timestamp
    m = re.match(r'^([^_]+)_(.+)$', parent)
    if m:
        second = m.group(2)
        if not re.search(r'_\d{10,}$', second):
            return m.group(1), second

    # Fallback: parent looks like a filename pattern (e.g. 5A0737_090R06_202603310955)
    m = re.match(r'^([^_]+)', parent)
    lot_no = m.group(1) if m else parent
    # Use filename's first segment as WaferID fallback
    fname = os.path.splitext(parts[-1])[0]
    fm = re.match(r'^([^_]+)_', fname)
    wafer_id = fm.group(1) if fm else ''
    return lot_no, wafer_id


def has_rh_sheet(filepath):
    try:
        wb = load_workbook(filepath, read_only=True, keep_links=False)
        has = 'RH' in wb.sheetnames
        wb.close()
        return has
    except Exception:
        return False


def compute_bj_percent(resistances):
    diffs = np.abs(np.diff(resistances))
    if len(diffs) == 0:
        return 0.0
    max_diff = np.max(diffs)
    r_range = np.max(resistances) - np.min(resistances)
    if r_range == 0:
        return 0.0
    return max_diff / r_range * 100.0


def process_file(filepath):
    """读取一个文件，返回 [(lot, wafer, chip, channel, bj_100, bj_200), ...]"""
    lot_no, wafer_id = parse_lot_wafer(filepath)

    try:
        df = pd.read_excel(filepath, sheet_name='RH', engine='openpyxl')
    except Exception:
        return None, None

    df.columns = [str(c).strip() for c in df.columns]

    # ChipID from Part ID
    chip_ids = df['Part ID'].dropna().unique()
    if len(chip_ids) == 0:
        return None, None
    chip_id = str(chip_ids[0]).strip()

    # Filter: only 'R' in ChipID
    if 'R' not in chip_id:
        return None, None

    df = df.sort_values(['Hd', 'Field(Gs)'])

    rows = []
    for ch in CHANNEL_ORDER:
        bj_100 = np.nan
        bj_200 = np.nan

        # ±200Gs 区间
        sub200 = df[(df['Hd'] == ch) & (df['Field(Gs)'] >= -200) & (df['Field(Gs)'] <= 200)]
        if len(sub200) >= 2:
            fw200 = sub200['Forward(Ohm)'].to_numpy(dtype=float)
            rev200 = sub200['Reverse(Ohm)'].to_numpy(dtype=float)
            bj_200 = max(compute_bj_percent(fw200), compute_bj_percent(rev200))
            bj_200 = round(bj_200, 4) if not np.isnan(bj_200) else np.nan

            # ±100Gs 区间：分子 = ±100Gs 相邻点差值最大值；分母 = ±200Gs 极差
            sub100 = df[(df['Hd'] == ch) & (df['Field(Gs)'] >= -100) & (df['Field(Gs)'] <= 100)]
            if len(sub100) >= 2:
                fw100 = sub100['Forward(Ohm)'].to_numpy(dtype=float)
                rev100 = sub100['Reverse(Ohm)'].to_numpy(dtype=float)
                # 分子: ±100Gs 区间相邻点差值最大值
                max_diff_100 = max(
                    np.max(np.abs(np.diff(fw100))) if len(fw100) >= 2 else 0.0,
                    np.max(np.abs(np.diff(rev100))) if len(rev100) >= 2 else 0.0,
                )
                # 分母: ±200Gs 区间极差 (Forward / Reverse 各算，取最大)
                range_fw200 = np.max(fw200) - np.min(fw200)
                range_rev200 = np.max(rev200) - np.min(rev200)
                r_range_200 = max(range_fw200, range_rev200)
                if r_range_200 > 0:
                    bj_100 = round(max_diff_100 / r_range_200 * 100.0, 4)
                else:
                    bj_100 = np.nan
            else:
                bj_100 = np.nan
        else:
            bj_200 = np.nan
            bj_100 = np.nan

        rows.append({
            'LotNo': lot_no,
            'WaferID': wafer_id,
            'ChipID': chip_id,
            'ChannelID': ch,
            'BJ%_100Gs': bj_100,
            'BJ%_200Gs': bj_200,
        })

    meta = {'LotNo': lot_no, 'WaferID': wafer_id, 'ChipID': chip_id,
            'File': os.path.basename(filepath)}
    return rows, meta


# ── HTML 生成 ──────────────────────────────────────────────

def generate_html(df_raw, out_path):
    """生成交互式 HTML"""
    # Prepare data for JSON
    records = df_raw.to_dict(orient='records')
    # Replace NaN with null for JSON
    for r in records:
        for k in ('BJ%_100Gs', 'BJ%_200Gs'):
            v = r[k]
            if pd.isna(v):
                r[k] = None

    lot_wafer = df_raw[['LotNo', 'WaferID']].drop_duplicates().sort_values(['LotNo', 'WaferID'])
    lot_wafer_list = lot_wafer.to_dict(orient='records')

    data_json = json.dumps(records, ensure_ascii=False)
    lotwafer_json = json.dumps(lot_wafer_list, ensure_ascii=False)
    channels_json = json.dumps(CHANNEL_ORDER)

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>RH BJ% 良率分析</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; padding: 16px; }}
  h1 {{ font-size: 22px; margin-bottom: 12px; color: #1a1a2e; }}
  .card {{ background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); padding: 18px 22px; margin-bottom: 16px; }}
  .card h2 {{ font-size: 16px; margin-bottom: 12px; color: #16213e; border-left: 4px solid #0f3460; padding-left: 10px; }}
  .controls {{ display: flex; flex-wrap: wrap; gap: 16px; align-items: center; margin-bottom: 10px; }}
  .controls label {{ font-size: 13px; font-weight: 600; color: #333; }}
  select {{ padding: 5px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; min-width: 160px; }}
  input[type=range] {{ width: 300px; vertical-align: middle; }}
  .val-display {{ font-weight: bold; color: #e94560; font-size: 15px; min-width: 50px; display: inline-block; }}
  .plot-row {{ display: flex; gap: 16px; flex-wrap: wrap; }}
  .plot-box {{ flex: 1; min-width: 500px; }}
  .stats-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  .stats-table th {{ background: #0f3460; color: #fff; padding: 6px 8px; text-align: center; }}
  .stats-table td {{ border: 1px solid #ddd; padding: 4px 8px; text-align: center; }}
  .stats-table tr:nth-child(even) {{ background: #f8f9fa; }}
  .summary-grid {{ display: flex; gap: 30px; flex-wrap: wrap; }}
  .summary-item {{ text-align: center; }}
  .summary-item .num {{ font-size: 28px; font-weight: bold; color: #0f3460; }}
  .summary-item .lbl {{ font-size: 12px; color: #666; }}
  .tab-bar {{ display: flex; gap: 4px; margin-bottom: 0; }}
  .tab-bar button {{ padding: 8px 20px; border: 1px solid #ccc; border-bottom: none; background: #e9ecef; cursor: pointer; font-size: 13px; border-radius: 4px 4px 0 0; }}
  .tab-bar button.active {{ background: #fff; border-color: #0f3460; font-weight: bold; color: #0f3460; }}
  .tab-content {{ display: none; }}
  .tab-content.active {{ display: block; }}
</style>
</head>
<body>
<h1>RH BJ% 良率分析仪表板</h1>

<div class="card">
  <h2>数据筛选</h2>
  <div class="controls">
    <label>LotNo:</label>
    <select id="selLot" onchange="updateWaferDropdown()"></select>
    <label>WaferID:</label>
    <select id="selWafer" onchange="updatePlots()"></select>
    <label style="margin-left:30px;">BJ% 上限:</label>
    <input type="range" id="ulSlider" min="1" max="20" step="0.5" value="5" oninput="onULChange()">
    <span class="val-display" id="ulVal">5.0%</span>
  </div>
</div>

<div class="card">
  <h2>BJ% 分布对比 (±100Gs vs ±200Gs)</h2>
  <div class="plot-row">
    <div class="plot-box"><div id="boxPlot" style="height:420px;"></div></div>
    <div class="plot-box"><div id="histPlot" style="height:420px;"></div></div>
  </div>
</div>

<div class="card">
  <h2>BJ%_100Gs vs BJ%_200Gs 散点图</h2>
  <div id="scatterPlot" style="width:1400px; height:760px; margin:0 auto; overflow-x:auto;"></div>
</div>

<div class="card">
  <h2>良率统计 (BJ% 上限 = <span id="ulLabel">5.0%</span>)</h2>
  <div class="summary-grid" id="summaryGrid"></div>
  <div class="tab-bar" style="margin-top:16px;">
    <button class="active" onclick="switchTab('ch', this)">按通道良率</button>
    <button onclick="switchTab('wafer', this)">按晶圆Chip良率</button>
  </div>
  <div class="tab-content active" id="tab-ch">
    <div id="chYieldBar" style="height:350px; margin-top:10px;"></div>
    <table class="stats-table" id="chTable" style="margin-top:10px;"></table>
  </div>
  <div class="tab-content" id="tab-wafer">
    <div id="waferYieldBar" style="height:350px; margin-top:10px;"></div>
    <table class="stats-table" id="waferTable" style="margin-top:10px;"></table>
  </div>
</div>

<script>
const ALL_DATA = {data_json};
const LOT_WAFER = {lotwafer_json};
const CHANNELS = {channels_json};

// ── Populate LotNo dropdown ──
const lots = [...new Set(LOT_WAFER.map(x => x.LotNo))].sort();
const selLot = document.getElementById('selLot');
lots.forEach(l => {{
  const opt = document.createElement('option');
  opt.value = l; opt.text = l;
  selLot.appendChild(opt);
}});

function getWafersForLot(lot) {{
  return LOT_WAFER.filter(x => x.LotNo === lot).map(x => x.WaferID).sort();
}}

function updateWaferDropdown() {{
  const lot = selLot.value;
  const selWafer = document.getElementById('selWafer');
  selWafer.innerHTML = '';
  getWafersForLot(lot).forEach(w => {{
    const opt = document.createElement('option');
    opt.value = w; opt.text = w;
    selWafer.appendChild(opt);
  }});
  updatePlots();
}}

function getCurrentData() {{
  const lot = document.getElementById('selLot').value;
  const wafer = document.getElementById('selWafer').value;
  return ALL_DATA.filter(r => r.LotNo === lot && r.WaferID === wafer);
}}

function updatePlots() {{
  const data = getCurrentData();
  drawBoxPlot(data);
  drawHistPlot(data);
  drawScatterPlot(data);
  updateYield();
}}

function drawBoxPlot(data) {{
  const v100 = data.filter(r => r['BJ%_100Gs'] !== null).map(r => r['BJ%_100Gs']);
  const v200 = data.filter(r => r['BJ%_200Gs'] !== null).map(r => r['BJ%_200Gs']);
  const traces = [];
  if (v100.length) traces.push({{y: v100, type:'box', name:'±100Gs', boxmean:'sd', marker:{{color:'#1f77b4'}}}});
  if (v200.length) traces.push({{y: v200, type:'box', name:'±200Gs', boxmean:'sd', marker:{{color:'#ff7f0e'}}}});
  Plotly.react('boxPlot', traces, {{
    title: 'BJ% Box Plot',
    yaxis:{{title:'BJ%'}},
    margin:{{l:50,r:30,t:40,b:40}},
    showlegend:true,
  }});
}}

function drawHistPlot(data) {{
  const v100 = data.filter(r => r['BJ%_100Gs'] !== null).map(r => r['BJ%_100Gs']);
  const v200 = data.filter(r => r['BJ%_200Gs'] !== null).map(r => r['BJ%_200Gs']);
  const traces = [];
  if (v100.length) traces.push({{x: v100, type:'histogram', name:'±100Gs', opacity:0.7, marker:{{color:'#1f77b4'}}, nbinsx:30}});
  if (v200.length) traces.push({{x: v200, type:'histogram', name:'±200Gs', opacity:0.7, marker:{{color:'#ff7f0e'}}, nbinsx:30}});
  Plotly.react('histPlot', traces, {{
    title: 'BJ% Distribution',
    xaxis:{{title:'BJ%'}},
    yaxis:{{title:'Count'}},
    barmode:'overlay',
    margin:{{l:50,r:30,t:40,b:40}},
    bargap:0.05,
  }});
}}

function onULChange() {{
  drawScatterPlot(getCurrentData());
  updateYield();
}}

function drawScatterPlot(data) {{
  const pairs = data.filter(r => r['BJ%_100Gs'] !== null && r['BJ%_200Gs'] !== null);
  const x = pairs.map(r => r['BJ%_100Gs']);
  const y = pairs.map(r => r['BJ%_200Gs']);
  const n = x.length;

  // Read UL from slider
  const ul = parseFloat(document.getElementById('ulSlider').value);

  // Static color palette
  const PALETTE = ['#1f77b4','#ff7f0e','#2ca02c','#d62728','#9467bd','#8c564b','#e377c2','#7f7f7f','#bcbd22','#17becf','#aec7e8','#ffbb78','#98df8a','#ff9896','#c5b0d5','#c49c94','#f7b6d2','#c7c7c7','#dbdb8d','#9edae5','#393b79','#637939','#8c6d31','#843c39','#7b4173','#bd9e89','#5254a3','#8ca252','#bd9e39','#e7969c','#de9ed6','#31a354','#ad494a','#756bb1'];

  // Group by channel
  const chGroups = {{}};
  pairs.forEach((r, i) => {{
    const ch = r.ChannelID;
    if (!chGroups[ch]) chGroups[ch] = {{x: [], y: []}};
    chGroups[ch].x.push(x[i]);
    chGroups[ch].y.push(y[i]);
  }});

  // Axis range for main plot
  const allVals = x.concat(y);
  const maxVal = Math.max(...allVals, 1);
  const axisMax = Math.ceil(maxVal * 1.05);
  const axisMin = 0;

  // Robust Cauchy fit
  let fitSlope = null, fitIntercept = null, fitR2 = null;
  if (n >= 2) {{
    const meanX0 = x.reduce((s,v) => s+v, 0) / n;
    const meanY0 = y.reduce((s,v) => s+v, 0) / n;
    let num0 = 0, den0 = 0;
    for (let i = 0; i < n; i++) {{ num0 += (x[i] - meanX0) * (y[i] - meanY0); den0 += (x[i] - meanX0) ** 2; }}
    fitSlope = den0 === 0 ? 0 : num0 / den0;
    fitIntercept = meanY0 - fitSlope * meanX0;
    const c = 2.385;
    for (let iter = 0; iter < 100; iter++) {{
      const residuals = [];
      for (let i = 0; i < n; i++) residuals.push(y[i] - (fitSlope * x[i] + fitIntercept));
      const absRes = residuals.map(Math.abs).sort((a,b) => a - b);
      const medRes = absRes[Math.floor(n / 2)];
      const mad = Math.max(medRes * 1.4826, 1e-8);
      const scale = c * mad;
      const weights = residuals.map(r => 1 / (1 + (r / scale) ** 2));
      let sw = 0, swx = 0, swy = 0, swxx = 0, swxy = 0;
      for (let i = 0; i < n; i++) {{ const w = weights[i]; sw += w; swx += w * x[i]; swy += w * y[i]; swxx += w * x[i] ** 2; swxy += w * x[i] * y[i]; }}
      const newSlope = (sw * swxy - swx * swy) / (sw * swxx - swx * swx + 1e-12);
      const newIntercept = (swy - newSlope * swx) / sw;
      if (Math.abs(newSlope - fitSlope) < 1e-8 && Math.abs(newIntercept - fitIntercept) < 1e-8) {{ fitSlope = newSlope; fitIntercept = newIntercept; break; }}
      fitSlope = newSlope; fitIntercept = newIntercept;
    }}
    const meanY = y.reduce((s,v) => s+v, 0) / n;
    let ssRes = 0, ssTot = 0;
    for (let i = 0; i < n; i++) {{ ssRes += (y[i] - (fitSlope * x[i] + fitIntercept)) ** 2; ssTot += (y[i] - meanY) ** 2; }}
    fitR2 = ssTot === 0 ? 1 : 1 - ssRes / ssTot;
  }}

  // Adjusted UL for ±100Gs
  let adjustedUL = ul;
  if (fitSlope !== null && fitSlope !== 0) {{
    adjustedUL = (ul - fitIntercept) / fitSlope;
  }}

  // Subplot axis ranges
  const subXMax = Math.max(adjustedUL * 1.15, 1);
  const subYMax = Math.max(ul * 1.15, 1);

  // ── Build traces ──
  const traces = [];

  // Main plot: scatter (with legend)
  Object.keys(chGroups).forEach((ch, idx) => {{
    traces.push({{
      x: chGroups[ch].x, y: chGroups[ch].y,
      mode: 'markers', type: 'scatter', name: ch,
      marker: {{size: 5, opacity: 0.7, color: PALETTE[idx % PALETTE.length]}},
      hoverinfo: 'skip',
    }});
  }});
  // Main plot: diagonal
  traces.push({{
    x: [axisMin, axisMax], y: [axisMin, axisMax],
    mode: 'lines', type: 'scatter', name: 'y = x',
    line: {{color: '#888', width: 1.5, dash: 'dash'}},
    hoverinfo: 'skip',
  }});
  // Main plot: fit line
  if (fitSlope !== null) {{
    traces.push({{
      x: [axisMin, axisMax], y: [fitSlope * axisMin + fitIntercept, fitSlope * axisMax + fitIntercept],
      mode: 'lines', type: 'scatter',
      name: `Robust Fit (Cauchy): y=${{fitSlope.toFixed(4)}}x+${{fitIntercept.toFixed(4)}} (R²=${{fitR2.toFixed(4)}})`,
      line: {{color: '#e94560', width: 2}},
      hoverinfo: 'skip',
    }});
  }}

  // Subplot: scatter (no legend)
  Object.keys(chGroups).forEach((ch, idx) => {{
    traces.push({{
      x: chGroups[ch].x, y: chGroups[ch].y,
      mode: 'markers', type: 'scatter', name: ch,
      xaxis: 'x2', yaxis: 'y2',
      showlegend: false,
      marker: {{size: 4, opacity: 0.6, color: PALETTE[idx % PALETTE.length]}},
      hoverinfo: 'skip',
    }});
  }});
  // Subplot: diagonal
  traces.push({{
    x: [0, Math.min(subXMax, subYMax)], y: [0, Math.min(subXMax, subYMax)],
    mode: 'lines', type: 'scatter', name: 'y = x',
    xaxis: 'x2', yaxis: 'y2',
    showlegend: false,
    line: {{color: '#888', width: 1.5, dash: 'dash'}},
    hoverinfo: 'skip',
  }});
  // Subplot: fit line
  if (fitSlope !== null) {{
    traces.push({{
      x: [0, subXMax], y: [fitIntercept, fitSlope * subXMax + fitIntercept],
      mode: 'lines', type: 'scatter',
      xaxis: 'x2', yaxis: 'y2',
      showlegend: false,
      line: {{color: '#e94560', width: 2}},
      hoverinfo: 'skip',
    }});
  }}
  // Subplot: UL horizontal line
  traces.push({{
    x: [0, subXMax], y: [ul, ul],
    mode: 'lines', type: 'scatter',
    xaxis: 'x2', yaxis: 'y2',
    showlegend: false,
    line: {{color: '#ff7f0e', width: 1, dash: 'dot'}},
    hoverinfo: 'skip',
  }});
  // Subplot: adjustedUL vertical line
  traces.push({{
    x: [adjustedUL, adjustedUL], y: [0, subYMax],
    mode: 'lines', type: 'scatter',
    xaxis: 'x2', yaxis: 'y2',
    showlegend: false,
    line: {{color: '#2ca02c', width: 1, dash: 'dot'}},
    hoverinfo: 'skip',
  }});

  // Annotations
  const annotations = [];
  // Main plot fit equation
  if (fitSlope !== null) {{
    annotations.push({{
      x: axisMax * 0.95, y: axisMin + (axisMax - axisMin) * 0.05,
      xref: 'x', yref: 'y',
      text: `y = ${{fitSlope.toFixed(4)}}x ${{fitIntercept >= 0 ? '+' : '-'}} ${{Math.abs(fitIntercept).toFixed(4)}}<br>R² = ${{fitR2.toFixed(4)}}`,
      showarrow: false, font: {{size: 13, color: '#e94560'}},
      bgcolor: 'rgba(255,255,255,0.85)', bordercolor: '#e94560', borderwidth: 1, borderpad: 6, align: 'left',
    }});
  }}
  // Subplot labels
  annotations.push({{
    x: 0.5, y: 1.035, xref: 'x2', yref: 'y2 domain',
    text: `Zoomed: X≤${{adjustedUL.toFixed(2)}}% (AdjUL), Y≤${{ul.toFixed(1)}}% (UL)`,
    showarrow: false, font: {{size: 12, color: '#333'}},
  }});
  // Subplot: UL label
  annotations.push({{
    x: subXMax * 0.98, y: ul, xref: 'x2', yref: 'y2',
    text: `UL=${{ul.toFixed(1)}}%`, showarrow: false,
    font: {{size: 10, color: '#ff7f0e'}}, xanchor: 'right', yanchor: 'bottom',
  }});
  // Subplot: adjustedUL label
  annotations.push({{
    x: adjustedUL, y: subYMax * 0.98, xref: 'x2', yref: 'y2',
    text: `AdjUL=${{adjustedUL.toFixed(2)}}%`, showarrow: false,
    font: {{size: 10, color: '#2ca02c'}}, xanchor: 'left', yanchor: 'top',
    textangle: 90,
  }});

  // Layout: two 1:1 subplots side by side
  // Canvas 1400x760; each plot area 580x580px
  // Left: x=[0.043, 0.457], y=[0.079, 0.842]
  // Right: x=[0.543, 0.957], y=[0.079, 0.842]
  const dtickMain = Math.ceil(axisMax / 10);
  const dtickSub = Math.max(Math.ceil(subXMax / 5), 0.5);
  const dtickSubY = Math.max(Math.ceil(subYMax / 5), 0.5);

  Plotly.react('scatterPlot', traces, {{
    title: 'BJ%_100Gs vs BJ%_200Gs',
    xaxis: {{title: 'BJ%_100Gs', range: [axisMin, axisMax], gridcolor: '#eee', zeroline: true, fixedrange: true, dtick: dtickMain, domain: [0.043, 0.457]}},
    yaxis: {{title: 'BJ%_200Gs', range: [axisMin, axisMax], gridcolor: '#eee', zeroline: true, fixedrange: true, dtick: dtickMain, domain: [0.079, 0.842]}},
    xaxis2: {{title: 'BJ%_100Gs', range: [0, subXMax], gridcolor: '#eee', zeroline: true, fixedrange: true, dtick: dtickSub, domain: [0.543, 0.957]}},
    yaxis2: {{title: 'BJ%_200Gs', range: [0, subYMax], gridcolor: '#eee', zeroline: true, fixedrange: true, dtick: dtickSubY, domain: [0.079, 0.842]}},
    margin: {{l:50, r:30, t:80, b:50}},
    showlegend: false,
    annotations: annotations,
    hovermode: false,
    width: 1400, height: 760,
  }}, {{
    staticPlot: true,
    displayModeBar: false,
  }});
}}

function computeRobustCauchyFit(x, y) {{
  const n = x.length;
  if (n < 2) return null;
  const meanX = x.reduce((s,v)=>s+v,0)/n;
  const meanY = y.reduce((s,v)=>s+v,0)/n;
  let num=0, den=0;
  for (let i=0;i<n;i++) {{ num += (x[i]-meanX)*(y[i]-meanY); den += (x[i]-meanX)**2; }}
  let slope = den===0 ? 0 : num/den;
  let intercept = meanY - slope*meanX;
  const c = 2.385;
  for (let iter=0; iter<100; iter++) {{
    const residuals = [];
    for (let i=0;i<n;i++) residuals.push(y[i]-(slope*x[i]+intercept));
    const absRes = residuals.map(Math.abs).sort((a,b)=>a-b);
    const medRes = absRes[Math.floor(n/2)];
    const mad = Math.max(medRes*1.4826, 1e-8);
    const scale = c*mad;
    const weights = residuals.map(r => 1/(1+(r/scale)**2));
    let sw=0,swx=0,swy=0,swxx=0,swxy=0;
    for (let i=0;i<n;i++) {{
      const w=weights[i]; sw+=w; swx+=w*x[i]; swy+=w*y[i]; swxx+=w*x[i]**2; swxy+=w*x[i]*y[i];
    }}
    const newSlope = (sw*swxy-swx*swy)/(sw*swxx-swx*swx+1e-12);
    const newIntercept = (swy-newSlope*swx)/sw;
    if (Math.abs(newSlope-slope)<1e-8 && Math.abs(newIntercept-intercept)<1e-8) {{
      slope=newSlope; intercept=newIntercept; break;
    }}
    slope=newSlope; intercept=newIntercept;
  }}
  const my = y.reduce((s,v)=>s+v,0)/n;
  let ssRes=0, ssTot=0;
  for (let i=0;i<n;i++) {{
    ssRes += (y[i]-(slope*x[i]+intercept))**2;
    ssTot += (y[i]-my)**2;
  }}
  const r2 = ssTot===0 ? 1 : 1-ssRes/ssTot;
  return {{slope, intercept, r2}};
}}

function updateYield() {{
  const ul = parseFloat(document.getElementById('ulSlider').value);
  document.getElementById('ulVal').textContent = ul.toFixed(1) + '%';
  document.getElementById('ulLabel').textContent = ul.toFixed(1) + '%';

  const data = getCurrentData();
  const nChips = new Set(data.map(r => r.ChipID)).size;
  const channels = CHANNELS;

  // Compute robust fit on current wafer data
  const pairs = data.filter(r => r['BJ%_100Gs'] !== null && r['BJ%_200Gs'] !== null);
  const fx = pairs.map(r => r['BJ%_100Gs']);
  const fy = pairs.map(r => r['BJ%_200Gs']);
  const fit = computeRobustCauchyFit(fx, fy);
  // Adjusted UL for ±100Gs: solve x when y=ul  =>  x = (ul - intercept) / slope
  let adjustedUL = ul;
  if (fit && fit.slope !== 0) {{
    adjustedUL = (ul - fit.intercept) / fit.slope;
  }}

  // Channel yield
  const chPass100 = {{}}, chPass200 = {{}}, chPassAdj = {{}}, chTotal = {{}};
  channels.forEach(ch => {{ chPass100[ch]=0; chPass200[ch]=0; chPassAdj[ch]=0; chTotal[ch]=0; }});
  data.forEach(r => {{
    if (chTotal[r.ChannelID] !== undefined) {{
      chTotal[r.ChannelID]++;
      if (r['BJ%_100Gs'] !== null && r['BJ%_100Gs'] <= ul) chPass100[r.ChannelID]++;
      if (r['BJ%_200Gs'] !== null && r['BJ%_200Gs'] <= ul) chPass200[r.ChannelID]++;
      if (r['BJ%_100Gs'] !== null && r['BJ%_100Gs'] <= adjustedUL) chPassAdj[r.ChannelID]++;
    }}
  }});

  // Chip yield: all 34 channels pass
  const chipMap = {{}};
  data.forEach(r => {{
    if (!chipMap[r.ChipID]) chipMap[r.ChipID] = {{pass100:0, pass200:0, passAdj:0, total:0}};
    chipMap[r.ChipID].total++;
    if (r['BJ%_100Gs'] !== null && r['BJ%_100Gs'] <= ul) chipMap[r.ChipID].pass100++;
    if (r['BJ%_200Gs'] !== null && r['BJ%_200Gs'] <= ul) chipMap[r.ChipID].pass200++;
    if (r['BJ%_100Gs'] !== null && r['BJ%_100Gs'] <= adjustedUL) chipMap[r.ChipID].passAdj++;
  }});
  let chipPass100 = 0, chipPass200 = 0, chipPassAdj = 0;
  Object.values(chipMap).forEach(c => {{
    if (c.pass100 === c.total) chipPass100++;
    if (c.pass200 === c.total) chipPass200++;
    if (c.passAdj === c.total) chipPassAdj++;
  }});

  // Summary cards
  const avgCh100 = channels.reduce((s,ch) => s + (chTotal[ch] ? chPass100[ch]/chTotal[ch]*100 : 0), 0) / channels.length;
  const avgCh200 = channels.reduce((s,ch) => s + (chTotal[ch] ? chPass200[ch]/chTotal[ch]*100 : 0), 0) / channels.length;
  const avgChAdj = channels.reduce((s,ch) => s + (chTotal[ch] ? chPassAdj[ch]/chTotal[ch]*100 : 0), 0) / channels.length;
  const adjStr = fit ? adjustedUL.toFixed(2) + '%' : 'N/A';
  document.getElementById('summaryGrid').innerHTML = `
    <div class="summary-item"><div class="num">${{nChips}}</div><div class="lbl">Total Chips</div></div>
    <div class="summary-item"><div class="num">${{chipPass100}}/${{nChips}}</div><div class="lbl">±100Gs Chip Yield (${{(chipPass100/nChips*100).toFixed(1)}}%)</div></div>
    <div class="summary-item"><div class="num">${{chipPass200}}/${{nChips}}</div><div class="lbl">±200Gs Chip Yield (${{(chipPass200/nChips*100).toFixed(1)}}%)</div></div>
    <div class="summary-item"><div class="num">${{chipPassAdj}}/${{nChips}}</div><div class="lbl">Adjusted200Gs Chip Yield (${{(chipPassAdj/nChips*100).toFixed(1)}}%)<br>Adj UL=${{adjStr}}</div></div>
    <div class="summary-item"><div class="num">${{avgCh100.toFixed(1)}}%</div><div class="lbl">±100Gs Avg Ch Yield</div></div>
    <div class="summary-item"><div class="num">${{avgCh200.toFixed(1)}}%</div><div class="lbl">±200Gs Avg Ch Yield</div></div>
    <div class="summary-item"><div class="num">${{avgChAdj.toFixed(1)}}%</div><div class="lbl">Adjusted200Gs Avg Ch Yield</div></div>
  `;

  // Channel yield bar chart
  const y100 = channels.map(ch => chTotal[ch] ? chPass100[ch]/chTotal[ch]*100 : 0);
  const y200 = channels.map(ch => chTotal[ch] ? chPass200[ch]/chTotal[ch]*100 : 0);
  const yAdj = channels.map(ch => chTotal[ch] ? chPassAdj[ch]/chTotal[ch]*100 : 0);
  Plotly.react('chYieldBar', [
    {{x: channels, y: y100, type:'bar', name:'±100Gs', marker:{{color:'#1f77b4'}}}},
    {{x: channels, y: y200, type:'bar', name:'±200Gs', marker:{{color:'#ff7f0e'}}}},
    {{x: channels, y: yAdj, type:'bar', name:'Adjusted200Gs', marker:{{color:'#2ca02c'}}}},
  ], {{
    title: 'Channel Yield % (UL=' + ul.toFixed(1) + '%, AdjUL=' + adjStr + ')',
    yaxis:{{title:'Yield%', range:[0,105]}},
    xaxis:{{title:'Channel', tickangle:-45}},
    barmode:'group',
    margin:{{l:50,r:30,t:50,b:80}},
  }});

  // Channel table
  let html = '<tr><th>Channel</th><th>±100Gs<br>P/T</th><th>±100Gs<br>Yield%</th><th>±200Gs<br>P/T</th><th>±200Gs<br>Yield%</th><th>Adj200Gs<br>P/T</th><th>Adj200Gs<br>Yield%</th></tr>';
  channels.forEach(ch => {{
    const t = chTotal[ch] || 0;
    const p1 = chPass100[ch], p2 = chPass200[ch], p3 = chPassAdj[ch];
    html += `<tr><td>${{ch}}</td><td>${{p1}}/${{t}}</td><td>${{t?(p1/t*100).toFixed(1):0}}</td><td>${{p2}}/${{t}}</td><td>${{t?(p2/t*100).toFixed(1):0}}</td><td>${{p3}}/${{t}}</td><td>${{t?(p3/t*100).toFixed(1):0}}</td></tr>`;
  }});
  document.getElementById('chTable').innerHTML = html;

  // Wafer-level: group all wafers for current lot
  const lot = document.getElementById('selLot').value;
  const lotData = ALL_DATA.filter(r => r.LotNo === lot);
  const waferChipMap = {{}};
  lotData.forEach(r => {{
    if (!waferChipMap[r.WaferID]) waferChipMap[r.WaferID] = {{}};
    if (!waferChipMap[r.WaferID][r.ChipID]) waferChipMap[r.WaferID][r.ChipID] = {{pass100:0, pass200:0, passAdj:0, total:0}};
    const c = waferChipMap[r.WaferID][r.ChipID];
    c.total++;
    if (r['BJ%_100Gs'] !== null && r['BJ%_100Gs'] <= ul) c.pass100++;
    if (r['BJ%_200Gs'] !== null && r['BJ%_200Gs'] <= ul) c.pass200++;
    if (r['BJ%_100Gs'] !== null && r['BJ%_100Gs'] <= adjustedUL) c.passAdj++;
  }});

  const wafers = Object.keys(waferChipMap).sort();
  const wYield100 = wafers.map(w => {{
    const chips = Object.values(waferChipMap[w]);
    const pass = chips.filter(c => c.pass100 === c.total).length;
    return chips.length ? pass / chips.length * 100 : 0;
  }});
  const wYield200 = wafers.map(w => {{
    const chips = Object.values(waferChipMap[w]);
    const pass = chips.filter(c => c.pass200 === c.total).length;
    return chips.length ? pass / chips.length * 100 : 0;
  }});
  const wYieldAdj = wafers.map(w => {{
    const chips = Object.values(waferChipMap[w]);
    const pass = chips.filter(c => c.passAdj === c.total).length;
    return chips.length ? pass / chips.length * 100 : 0;
  }});

  Plotly.react('waferYieldBar', [
    {{x: wafers, y: wYield100, type:'bar', name:'±100Gs', marker:{{color:'#1f77b4'}}}},
    {{x: wafers, y: wYield200, type:'bar', name:'±200Gs', marker:{{color:'#ff7f0e'}}}},
    {{x: wafers, y: wYieldAdj, type:'bar', name:'Adjusted200Gs', marker:{{color:'#2ca02c'}}}},
  ], {{
    title: 'Chip Yield % by Wafer (Lot=' + lot + ', UL=' + ul.toFixed(1) + '%, AdjUL=' + adjStr + ')',
    yaxis:{{title:'Chip Yield%', range:[0,105]}},
    xaxis:{{title:'WaferID', tickangle:-45}},
    barmode:'group',
    margin:{{l:50,r:30,t:50,b:80}},
  }});

  let whtml = '<tr><th>WaferID</th><th>Total</th><th>±100Gs<br>Pass</th><th>±100Gs<br>Yield%</th><th>±200Gs<br>Pass</th><th>±200Gs<br>Yield%</th><th>Adj200Gs<br>Pass</th><th>Adj200Gs<br>Yield%</th></tr>';
  wafers.forEach(w => {{
    const chips = Object.values(waferChipMap[w]);
    const t = chips.length;
    const p1 = chips.filter(c => c.pass100 === c.total).length;
    const p2 = chips.filter(c => c.pass200 === c.total).length;
    const p3 = chips.filter(c => c.passAdj === c.total).length;
    whtml += `<tr><td>${{w}}</td><td>${{t}}</td><td>${{p1}}</td><td>${{t?(p1/t*100).toFixed(1):0}}</td><td>${{p2}}</td><td>${{t?(p2/t*100).toFixed(1):0}}</td><td>${{p3}}</td><td>${{t?(p3/t*100).toFixed(1):0}}</td></tr>`;
  }});
  document.getElementById('waferTable').innerHTML = whtml;
}}

function switchTab(tab, btn) {{
  document.querySelectorAll('.tab-bar button').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
}}

// Init
updateWaferDropdown();
</script>
</body>
</html>
'''

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


# ── 主流程 ─────────────────────────────────────────────────

def main():
    xlsx_files = glob.glob(os.path.join(ROOT_DIR, '**', '*.xlsx'), recursive=True)
    xlsx_files = [f for f in xlsx_files if '~$' not in f and 'RH_BJ_Yield' not in f and 'analyze_rh' not in f]
    print(f'找到 {len(xlsx_files)} 个 xlsx 文件', flush=True)

    # Filter to files with RH sheet
    print('检查 RH sheet...', flush=True)
    valid_files = []
    for i, fpath in enumerate(xlsx_files, 1):
        if i % 500 == 0:
            print(f'  检查进度 {i}/{len(xlsx_files)}', flush=True)
        if has_rh_sheet(fpath):
            valid_files.append(fpath)
    print(f'其中 {len(valid_files)} 个文件包含 RH sheet', flush=True)

    all_rows = []
    checkpoint_path = os.path.join(OUT_DIR, 'checkpoint.csv')
    for i, fpath in enumerate(valid_files, 1):
        lot, wafer = parse_lot_wafer(fpath)
        if i % 100 == 0:
            print(f'[{i}/{len(valid_files)}] {lot} / {wafer}', flush=True)
        rows, meta = process_file(fpath)
        if rows:
            all_rows.extend(rows)
        if i % 500 == 0 and all_rows:
            try:
                pd.DataFrame(all_rows).to_csv(checkpoint_path, index=False, encoding='utf-8-sig')
                print(f'  Checkpoint: {len(all_rows)} rows', flush=True)
            except Exception:
                pass

    print(f'\n处理完成: {len(all_rows)} 行 (ChipID含R)', flush=True)

    df_raw = pd.DataFrame(all_rows)
    # Filter only 'R' in ChipID (double-check, already filtered in process_file)
    df_raw = df_raw[df_raw['ChipID'].str.contains('R', na=False)].reset_index(drop=True)

    # Save CSV
    csv_path = os.path.join(OUT_DIR, 'RH_BJ_RawData.csv')
    tmp_csv = csv_path + '.tmp'
    df_raw.to_csv(tmp_csv, index=False, encoding='utf-8-sig')
    try:
        if os.path.exists(csv_path):
            os.remove(csv_path)
        os.rename(tmp_csv, csv_path)
    except PermissionError:
        os.rename(tmp_csv, csv_path.replace('.csv', '_new.csv'))
        csv_path = csv_path.replace('.csv', '_new.csv')
    print(f'原始数据已保存: {csv_path}')

    # Generate HTML
    html_path = os.path.join(OUT_DIR, 'RH_BJ_Yield_Dashboard.html')
    generate_html(df_raw, html_path)
    print(f'交互式HTML已保存: {html_path}')

    # Clean checkpoint
    try:
        os.remove(checkpoint_path)
    except Exception:
        pass

    print(f'\n=== 汇总 ===')
    print(f'芯片数: {df_raw["ChipID"].nunique()}')
    print(f'晶圆数: {df_raw["WaferID"].nunique()}')
    print(f'Lot数: {df_raw["LotNo"].nunique()}')
    print(f'数据行数: {len(df_raw)}')
    print(f'\nLotNo / WaferID 列表:')
    print(df_raw[['LotNo','WaferID']].drop_duplicates().sort_values('LotNo').to_string(index=False))


if __name__ == '__main__':
    main()
