import os
import re
import json
import time
import pandas as pd
import xlrd
from datetime import datetime
from openpyxl import load_workbook

SOURCE_DIR = r'\\server\share\Raw data'  # 数据源目录, 按实际环境修改
BAR_DIR = os.path.join(SOURCE_DIR, 'Bar')
FPC_DIR = os.path.join(SOURCE_DIR, 'FPC')
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
PROGRESS_FILE_BAR = os.path.join(OUTPUT_DIR, 'merge_progress_bar.json')
PROGRESS_FILE_FPC = os.path.join(OUTPUT_DIR, 'merge_progress_fpc.json')
OUTPUT_FILE_BAR = os.path.join(OUTPUT_DIR, 'Merged_Result.xlsx')
OUTPUT_FILE_FPC = os.path.join(OUTPUT_DIR, 'Merged_Result(FPC).xlsx')

XLS_RESULT_SHEETS = ['Transverse IBias.200G Result', 'Transverse.TC Result']

UNIFIED_COLS = ['WaferID', 'LotNo', 'Hd', 'Test_Date', 'PartID',
                'R0', 'Rmin', 'Rmax', 'AYSM', 'BJMP%', 'MR_Ratio']

XLS_COL_MAP = {
    'Hd': 'Hd',
    'Date': 'Test_Date',
    'Part ID': 'PartID',
    'Resistance (Ohm)': 'R0',
    'Min (Ohm)': 'Rmin',
    'Max (Ohm)': 'Rmax',
    'Pk-Pk Asym (%)': 'AYSM',
    'Barkh. Jump (%)': 'BJMP%',
    'Delta R/R (%)': 'MR_Ratio',
}

XLSX_COL_MAP = {
    'Hd': 'Hd',
    'Test time': 'Test_Date',
    'Part ID': 'PartID',
    'R(Ohm)': 'R0',
    'R Min(Ohm)': 'Rmin',
    'R Max(Ohm)': 'Rmax',
    'Asymmetric': 'AYSM',
    'Barkh. Jump (%)': 'BJMP%',
    'MR Ratio(%)': 'MR_Ratio',
}


def extract_lotno(lot_str):
    """Extract the 6-char LotNo (first 2 digits + 1 alnum + last 3 digits) from the start."""
    m = re.match(r'^\d{2}\w\d{3}', lot_str)
    return m.group(0) if m else lot_str


def parse_dir_name(dirname):
    m1 = re.match(r'^(MD\w+)\s*\((\w+)\)$', dirname)
    if m1:
        return m1.group(1), extract_lotno(m1.group(2))
    m2 = re.match(r'^(.+)_(MD\w+)$', dirname)
    if m2:
        return m2.group(2), extract_lotno(m2.group(1))
    m3 = re.match(r'^(.+)\s+(MD\w+)$', dirname)
    if m3:
        return m3.group(2), extract_lotno(m3.group(1))
    return None, None


def load_progress(progress_file):
    if os.path.exists(progress_file):
        with open(progress_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'processed_files': []}


def save_progress(progress, progress_file):
    with open(progress_file, 'w', encoding='utf-8') as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)


def get_xls_sheet_names(filepath):
    try:
        wb = xlrd.open_workbook(filepath, on_demand=True)
        names = wb.sheet_names()
        wb.release_resources()
        return names
    except Exception:
        return []


def find_sheet(available, candidates):
    for c in candidates:
        if c in available:
            return c
    return None


def remap_columns(df, col_map):
    existing = {k: v for k, v in col_map.items() if k in df.columns}
    if not existing:
        return None
    result = df[list(existing.keys())].rename(columns=existing)
    return result


def collect_new_files(source_dir, processed):
    """Scan source dir and return list of (rel_path, filepath, wafer_id, lot_no) for new files only."""
    new_files = []
    if not os.path.isdir(source_dir):
        print(f'Warning: Source directory does not exist: {source_dir}')
        return new_files
    for dirname in sorted(os.listdir(source_dir)):
        dirpath = os.path.join(source_dir, dirname)
        if not os.path.isdir(dirpath):
            continue
        wafer_id, lot_no = parse_dir_name(dirname)
        if wafer_id is None:
            continue
        for filename in sorted(os.listdir(dirpath)):
            if not (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')):
                continue
            filepath = os.path.join(dirpath, filename)
            rel_path = os.path.relpath(filepath, source_dir)
            if rel_path in processed:
                continue
            new_files.append((rel_path, filepath, wafer_id, lot_no))
    return new_files


def append_to_output(new_df, output_file):
    """Append new rows to the output xlsx using openpyxl, avoiding full rewrite of existing data."""
    if not os.path.exists(output_file) or os.path.getsize(output_file) < 100:
        new_df.to_excel(output_file, index=False, engine='openpyxl')
        return

    wb = load_workbook(output_file)
    ws = wb.active
    header = [cell.value for cell in ws[1]]
    col_index = {name: i + 1 for i, name in enumerate(header)}

    start_row = ws.max_row + 1
    for _, row in new_df.iterrows():
        r = start_row
        for col_name in UNIFIED_COLS:
            if col_name in col_index:
                val = row.get(col_name)
                if pd.isna(val):
                    val = None
                ws.cell(row=r, column=col_index[col_name], value=val)
        start_row += 1
    wb.save(output_file)


def scan_and_merge(source_dir, output_file, progress_file, label):
    """Scan source_dir for Excel files and merge rows into output_file, tracking progress in progress_file."""
    print(f'\n--- Merging {label} ---')
    print(f'Source:  {source_dir}')
    print(f'Output:  {output_file}')

    t0 = time.time()
    progress = load_progress(progress_file)
    processed = set(progress['processed_files'])

    # Phase 1: scan for new files only (no file I/O, just directory listing)
    new_files = collect_new_files(source_dir, processed)
    print(f'Found {len(new_files)} new file(s).')

    if not new_files:
        print(f'No new files for {label}.')
        return

    # Phase 2: read only new source files
    new_frames = []
    new_count = 0
    batch_count = 0

    for rel_path, filepath, wafer_id, lot_no in new_files:
        ext = filepath.lower()

        try:
            df = None

            if ext.endswith('.xlsx'):
                raw = pd.read_excel(filepath, sheet_name='R', engine='openpyxl')
                df = remap_columns(raw, XLSX_COL_MAP)

            elif ext.endswith('.xls'):
                sheet_names = get_xls_sheet_names(filepath)
                sn = find_sheet(sheet_names, XLS_RESULT_SHEETS)
                if sn:
                    raw = pd.read_excel(filepath, sheet_name=sn, engine='xlrd')
                    df = remap_columns(raw, XLS_COL_MAP)

            if df is not None and not df.empty:
                df.insert(0, 'WaferID', wafer_id)
                df.insert(1, 'LotNo', lot_no)
                for c in UNIFIED_COLS:
                    if c not in df.columns:
                        df[c] = None
                df = df[UNIFIED_COLS]
                new_frames.append(df)

            new_count += 1
            batch_count += 1
            progress['processed_files'].append(rel_path)

            if batch_count % 50 == 0:
                elapsed = time.time() - t0
                print(f'  [{new_count} files, {elapsed:.0f}s] Appending batch...')
                if new_frames:
                    batch_df = pd.concat(new_frames, ignore_index=True)[UNIFIED_COLS]
                    append_to_output(batch_df, output_file)
                    new_frames = []
                progress['last_run'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                save_progress(progress, progress_file)

        except Exception as e:
            print(f'  SKIP {rel_path}: {e}')
            progress['processed_files'].append(rel_path)

    elapsed = time.time() - t0
    print(f'\nProcessed {new_count} new file(s) in {elapsed:.0f}s. Appending final batch...')

    if new_frames:
        final_df = pd.concat(new_frames, ignore_index=True)[UNIFIED_COLS]
        append_to_output(final_df, output_file)

    progress['last_run'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save_progress(progress, progress_file)

    print(f'{label} done! Total time: {time.time() - t0:.0f}s')


def migrate_old_progress():
    """If the old merge_progress.json exists but the new Bar/FPC progress files don't,
    migrate the old content into merge_progress_bar.json (since old files now live under Bar).
    This avoids re-processing all previously merged Bar rows and duplicating them in output."""
    old_progress = os.path.join(OUTPUT_DIR, 'merge_progress.json')
    if os.path.exists(old_progress):
        if not os.path.exists(PROGRESS_FILE_BAR):
            try:
                import shutil
                shutil.copyfile(old_progress, PROGRESS_FILE_BAR)
                print(f'Migrated old progress -> {os.path.basename(PROGRESS_FILE_BAR)}')
            except Exception as e:
                print(f'Warning: failed to migrate old progress: {e}')
        # else: bar progress already exists, leave it alone


if __name__ == '__main__':
    print('=== Excel Merge Tool ===')
    print(f'Base dir: {SOURCE_DIR}')
    print(f'Output to: {OUTPUT_DIR}')

    # Migrate legacy progress (prevents duplicate rows in Merged_Result.xlsx after path change)
    migrate_old_progress()

    # 1. Merge original Bar data -> Merged_Result.xlsx
    scan_and_merge(BAR_DIR, OUTPUT_FILE_BAR, PROGRESS_FILE_BAR, 'Bar')

    # 2. Merge new FPC data -> Merged_Result(FPC).xlsx
    scan_and_merge(FPC_DIR, OUTPUT_FILE_FPC, PROGRESS_FILE_FPC, 'FPC')

    print('\n=== All merge tasks finished ===')
