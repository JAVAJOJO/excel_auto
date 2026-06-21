"""
资产序列号查询脚本 (Pandas 版)
"""
import argparse
import logging
import pandas as pd
from pathlib import Path

# ===================== 配置区 =====================
ASSET_FILE = Path(r"C:\Users\Administrator\Desktop\py\VFS IT Inventory Guangzhou 2026.xlsx")
ASSET_SHEET = None               # None 表示读取所有 sheet
ASSET_MATCH_COL = "N"            # pandas 列字母，需转换为 0-based 索引或列名
ASSET_HEADER_ROW = 1             # pandas 默认第0行是标题，如果你的标题在第2行(1-index)，需指定 header=1

RECEIVE_FILE = Path(r"C:\Users\Administrator\Desktop\py\Asset ID Status_FAR.xlsx")
RECEIVE_SHEET = None
RECEIVE_MATCH_COL = "G"

SPARE_FILE = Path(r"C:\Users\Administrator\Desktop\py\China Spare Asset.xlsx")
SPARE_SHEET = "Raw Data"
SPARE_MATCH_COL = "M"
# =================================================

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("FastQuery")

def col_letter_to_index(letter: str) -> int:
    """将 Excel 列字母转换为 0-based 索引"""
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter) - 1

def query_with_pandas(file_path: Path, sheet_name, match_col_letter: str, serial: str, label: str, header_row: int):
    if not file_path.exists():
        logger.info(f"[{label}] 文件不存在: {file_path}")
        return

    # 读取所有 sheet 或指定 sheet，使用 header 指定标题行（1-index）
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
    except Exception as e:
        logger.error(f"无法打开文件 {file_path}: {e}")
        return

    sheets_to_read = xls.sheet_names if sheet_name is None else [sheet_name]
    match_col_idx = col_letter_to_index(match_col_letter)

    found = False
    for sheet in sheets_to_read:
        # 读取整个 sheet（如果数据量大，可增加 usecols=[match_col_idx] 加速）
        df = pd.read_excel(xls, sheet_name=sheet, header=header_row-1, dtype=str)
        # 确保列索引存在
        if match_col_idx >= len(df.columns):
            continue
        # 查找匹配行
        mask = df.iloc[:, match_col_idx].str.strip() == serial
        if mask.any():
            row_idx = mask.idxmax()  # 第一个匹配的行索引
            logger.info(f"\n{'='*40}")
            logger.info(f"[{label}] 文件: {file_path.name}")
            logger.info(f"  工作表: {sheet} | 行号: {row_idx + header_row + 1}")  # 实际Excel行号
            # 输出整行
            row = df.iloc[row_idx]
            for col_name, value in row.items():
                logger.info(f"    {col_name}: {value}")
            found = True
            break
    if not found:
        logger.info(f"[{label}] 未找到序列号 {serial}")

def main():
    parser = argparse.ArgumentParser(description="资产序列号极速查询 (Pandas)")
    parser.add_argument("serial", help="要查询的序列号")
    args = parser.parse_args()
    serial = args.serial.strip()
    logger.info(f"查询序列号: {serial}")

    query_with_pandas(ASSET_FILE, ASSET_SHEET, ASSET_MATCH_COL, serial, "资产表", ASSET_HEADER_ROW)  # 标题第1行
    query_with_pandas(RECEIVE_FILE, RECEIVE_SHEET, RECEIVE_MATCH_COL, serial, "收货表", header_row=1)
    query_with_pandas(SPARE_FILE, SPARE_SHEET, SPARE_MATCH_COL, serial, "Spare表", header_row=1)

if __name__ == "__main__":
    main()