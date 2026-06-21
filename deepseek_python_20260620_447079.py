"""
多文件多字段资产查询（Pandas）
支持按字段名在多个 Excel 文件中搜索，也可全局搜索所有字段。
输出对齐易读。
"""

import argparse
import logging
from pathlib import Path
from typing import List, Dict
import pandas as pd
from openpyxl.utils import column_index_from_string

# ===================== 配置区 =====================
FILE_CONFIGS = {
    "asset": {
        "path": Path(r"C:\Users\itsupport_gz\OneDrive - VFSGlobal\China IT Asset - General\Inventory\IT Asset Tagging Tracker\South China\Guangzhou\VFS IT Inventory Guangzhou 2026.xlsx"),
        "sheet": "all",
        "header_row": 2
    },
    "receive": {
        "path": Path("2.xlsx"),
        "sheet": "all",
        "header_row": 1
    },
    "spare": {
        "path": Path("Spare.xlsx"),
        "sheet": "all",
        "header_row": 1
    }
}

FIELD_MAP = {
    "serial": [
        ("asset", "N"),
        ("receive", "G"),
        ("spare", "M")
    ],
    "fin_asset": [
        ("asset", "AF"),
    ],
    "global_asset": [
        ("asset", "AG"),
    ],
    "hostname": [
        ("asset", "C"),
        ("spare", "D")
    ]
}
# =================================================

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("MultiFileQuery")

def col_letter_to_idx(letter: str) -> int:
    return column_index_from_string(letter) - 1

def read_file(file_alias: str):
    cfg = FILE_CONFIGS.get(file_alias)
    if not cfg:
        logger.error(f"未知文件别名: {file_alias}")
        return []

    file_path = cfg["path"]
    if not file_path.exists():
        logger.warning(f"文件不存在，跳过: {file_path}")
        return []

    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
    except Exception as e:
        logger.error(f"无法打开文件 {file_path}: {e}")
        return []

    sheet_mode = cfg["sheet"]
    sheets = xls.sheet_names if sheet_mode.lower() == "all" else [sheet_mode]
    header_0based = cfg["header_row"] - 1

    data = []
    for sheet in sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header_0based, dtype=str)
            data.append((sheet, df, cfg["header_row"]))
        except Exception as e:
            logger.warning(f"读取工作表 [{sheet}] 失败: {e}")
    return data

def search_field(field: str, value: str) -> List[Dict]:
    results = []
    if field not in FIELD_MAP:
        logger.error(f"不支持的查询字段: {field}。可选: {list(FIELD_MAP.keys())}")
        return results

    for file_alias, col_letter in FIELD_MAP[field]:
        col_idx = col_letter_to_idx(col_letter)
        sheets_data = read_file(file_alias)
        for sheet_name, df, header_row in sheets_data:
            if col_idx >= len(df.columns):
                continue
            col_data = df.iloc[:, col_idx].astype(str).str.strip()
            mask = col_data == str(value).strip()
            if mask.any():
                for idx in mask[mask].index:
                    row = df.iloc[idx]
                    excel_row = idx + header_row + 1
                    results.append({
                        "file": file_alias,
                        "path": FILE_CONFIGS[file_alias]["path"].name,
                        "sheet": sheet_name,
                        "excel_row": excel_row,
                        "search_field": field,
                        "search_value": value,
                        "row_data": row
                    })
    return results

def search_all_fields(value: str) -> List[Dict]:
    results = []
    for field in FIELD_MAP:
        results.extend(search_field(field, value))
    return results

def print_results(results: List[Dict]):
    if not results:
        logger.info("未找到任何匹配。")
        return

    for i, res in enumerate(results, 1):
        row = res['row_data']
        items = []
        for col_name, cell_value in row.items():
            if pd.isna(cell_value):
                display_val = "(空)"
            else:
                display_val = str(cell_value).strip()
            items.append((str(col_name), display_val))

        max_key_width = max((len(k) for k, _ in items), default=8)
        max_key_width = max(max_key_width, 8)

        logger.info(f"\n{'='*60}")
        logger.info(f"  📁 文件 : {res['path']}  ({res['file']})")
        logger.info(f"  📋 工作表: {res['sheet']}    行号: {res['excel_row']}")
        logger.info(f"  🔍 匹配字段: {res['search_field']} = {res['search_value']}")
        logger.info(f"{'─'*60}")
        for key, val in items:
            logger.info(f"  {key:<{max_key_width}} : {val}")
        logger.info(f"{'─'*60}")

def main():
    parser = argparse.ArgumentParser(description="多文件多字段资产查询")
    parser.add_argument("value", help="要查找的值")
    parser.add_argument("-f", "--field", default="any",
                        help=f"字段名，可选: {list(FIELD_MAP.keys())}。默认 'any' 搜索所有字段")
    parser.add_argument("--file", help="限定仅在某个文件中搜索（文件别名）")
    args = parser.parse_args()

    value = args.value.strip()
    field = args.field.lower()

    # 如果指定了 --file，过滤 FIELD_MAP
    original_field_map = {k: v.copy() for k, v in FIELD_MAP.items()}
    if args.file:
        if args.file not in FILE_CONFIGS:
            logger.error(f"未知文件别名: {args.file}，可用: {list(FILE_CONFIGS.keys())}")
            return
        if field != "any":
            FIELD_MAP[field] = [(f, c) for f, c in FIELD_MAP.get(field, []) if f == args.file]
        else:
            for k in FIELD_MAP:
                FIELD_MAP[k] = [(f, c) for f, c in FIELD_MAP[k] if f == args.file]

    if field == "any":
        logger.info(f"搜索所有字段: '{value}'")
        results = search_all_fields(value)
    else:
        logger.info(f"按字段 '{field}' 搜索: '{value}'")
        results = search_field(field, value)

    # 还原 FIELD_MAP（如果被过滤了，下次运行不影响）
    FIELD_MAP.update(original_field_map)

    print_results(results)

if __name__ == "__main__":
    main()