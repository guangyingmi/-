#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按人名把汇总表拆成“一个Excel，多个以人名命名的Sheet”（保留样式，带进度条）
---------------------------------------------------------------------------
- 读取汇总Excel，按“销售员/姓名”等列分组
- 在一个新工作簿里，为每个人创建一个Sheet（Sheet名=人名，自动处理非法字符/重名/超长）
- 每个Sheet保留：表头样式、列宽、行高、边框、字体、对齐、冻结窗格、自动筛选
- 不做任何验证/报告

依赖：openpyxl>=3.1, tqdm>=4.60
用法示例：
    python excel_split_to_sheets.py -i "6月份利润.xlsx"
    python excel_split_to_sheets.py -i "6月份利润.xlsx" -s Sheet1 -c 销售员 -o 人员分Sheet.xlsx
"""
import argparse
import os
import re
import sys
import datetime as _dt
from collections import OrderedDict
from copy import copy
from typing import Optional, List, Dict

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from tqdm import tqdm

# 自动识别姓名列的候选关键词
DEFAULT_NAME_KEYS = ["销售员","姓名","员工","人员","负责人","Name","name"]


# ----------------- 通用工具 -----------------
def log(msg: str):
    ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def find_default_excel() -> Optional[str]:
    for name in os.listdir("."):
        if name.lower().endswith(".xlsx") and not name.startswith("~$"):
            if "按人拆分" in name:
                continue
            return name
    return None


def detect_sheet(wb, sheet):
    if sheet is None:
        return wb.worksheets[0]
    if isinstance(sheet, str):
        if sheet.isdigit():
            idx = int(sheet)
            return wb.worksheets[idx]
        if sheet in wb.sheetnames:
            return wb[sheet]
    if isinstance(sheet, int):
        return wb.worksheets[sheet]
    return wb.worksheets[0]


def detect_name_col(header_cells: List[str], manual: Optional[str]=None) -> str:
    if manual and manual in header_cells:
        return manual
    for key in DEFAULT_NAME_KEYS:
        if key in header_cells:
            return key
    for c in header_cells:
        if any(key in c for key in DEFAULT_NAME_KEYS):
            return c
    return header_cells[0]


def base_name(s) -> str:
    """把“张三 汇总”统一成“张三”"""
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"\s*汇总$", "", s)
    return s.strip()


def make_sheet_title(person: str, existing: set) -> str:
    """
    生成 Excel 合法且唯一的 sheet 名：
    - 禁止字符：: \ / ? * [ ]
    - 最长 31 字符
    - 如重名自动加 _2/_3 ...，并确保不超长
    """
    title = re.sub(r'[:\\/*?\[\]]', "_", (person or "未命名").strip() or "未命名")
    if len(title) > 31:
        title = title[:31]
    base = title
    i = 2
    while title in existing:
        suffix = f"_{i}"
        max_base_len = 31 - len(suffix)
        title = (base[:max_base_len] if max_base_len > 0 else base[:31-len(suffix)]) + suffix
        i += 1
    return title


# ----------------- 样式复制 -----------------
def copy_cell(src, dst):
    """复制单元格的值与样式（字体、对齐、边框、填充、数字格式、保护）。"""
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def copy_header_and_dimensions(src_ws, dst_ws, header_row=1):
    """复制列宽、表头样式、行高、冻结窗格。"""
    # 列宽
    for col_letter, dim in src_ws.column_dimensions.items():
        nd = dst_ws.column_dimensions[col_letter]
        nd.width = dim.width
        nd.hidden = dim.hidden
        nd.bestFit = dim.bestFit

    # 表头（源 header_row -> 目标第 1 行）
    for col in range(1, src_ws.max_column + 1):
        sc = src_ws.cell(row=header_row, column=col)
        dc = dst_ws.cell(row=1, column=col)
        copy_cell(sc, dc)

    # 行高
    try:
        dst_ws.row_dimensions[1].height = src_ws.row_dimensions[header_row].height
    except Exception:
        pass

    # 冻结窗格（常见为 A2）
    dst_ws.freeze_panes = src_ws.freeze_panes


def write_row_from_src(src_ws, dst_ws, src_row_idx, dst_row_idx):
    """按列逐格复制一整行（值+样式）。"""
    for col in range(1, src_ws.max_column + 1):
        sc = src_ws.cell(row=src_row_idx, column=col)
        dc = dst_ws.cell(row=dst_row_idx, column=col)
        copy_cell(sc, dc)


# ----------------- 主逻辑：一个Excel多Sheet -----------------
def split_to_sheets(in_path: str, sheet_sel, name_col_manual: Optional[str],
                    out_file: str, keep_empty: bool):
    log(f"输入文件：{in_path}")

    # 读源表（data_only=True 读取公式显示值）
    src_wb = load_workbook(in_path, data_only=True)
    src_ws = detect_sheet(src_wb, sheet_sel)
    log(f"工作表：{src_ws.title}")

    # 表头
    header = [str(c.value).strip() if c.value is not None else "" for c in next(src_ws.iter_rows(min_row=1, max_row=1))]
    if not header or all(not h for h in header):
        raise RuntimeError("无法读取表头（第 1 行为空）。")

    name_col = detect_name_col(header, name_col_manual)
    try:
        name_col_idx = header.index(name_col) + 1  # 1-based
    except ValueError:
        raise RuntimeError(f"未找到姓名列：{name_col}")
    log(f"使用姓名列：{name_col}（第 {name_col_idx} 列）")

    # 目标工作簿
    out_wb = Workbook()
    # 先删除默认空Sheet，后续按人创建
    default_ws = out_wb.active
    out_wb.remove(default_ws)

    # 人员 -> 目标Sheet
    books: Dict[str, object] = OrderedDict()
    existing_titles = set()
    header_row_idx = 1

    # 进度条：数据行数
    total_rows = src_ws.max_row - 1 if src_ws.max_row > 1 else 0
    pbar = tqdm(total=total_rows, desc="写入各人员Sheet", unit="行")

    # 遍历数据（从第2行开始）
    for r in range(2, src_ws.max_row + 1):
        person_raw = src_ws.cell(row=r, column=name_col_idx).value
        person = base_name(person_raw)
        if not person and not keep_empty:
            pbar.update(1)
            continue

        # 首次出现则创建 sheet 并复制表头设置
        if person not in books:
            title = make_sheet_title(person, existing_titles)
            existing_titles.add(title)
            dst_ws = out_wb.create_sheet(title=title)
            copy_header_and_dimensions(src_ws, dst_ws, header_row=header_row_idx)
            books[person] = dst_ws

        # 追加一行（带样式）
        dst_ws = books[person]
        dst_next_row = dst_ws.max_row + 1
        write_row_from_src(src_ws, dst_ws, r, dst_next_row)

        pbar.update(1)

    pbar.close()

    # 设置各sheet的自动筛选范围
    for person, ws in books.items():
        last_col_letter = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    # 保存
    out_wb.save(out_file)
    log(f"完成！共写入 {len(books)} 个人员Sheet -> {out_file}")


# ----------------- CLI -----------------
def main():
    ap = argparse.ArgumentParser(description="按人名把汇总表拆成一个Excel（多Sheet，保留样式）")
    ap.add_argument("-i","--input", help="输入 Excel 路径（默认自动扫描）")
    ap.add_argument("-s","--sheet", help="表名或索引（0基）", default=None)
    ap.add_argument("-c","--name-col", help="姓名列名（默认自动识别）")
    ap.add_argument("-o","--out-file", help="输出Excel文件路径（默认：按人分Sheet_时间戳.xlsx）")
    ap.add_argument("--keep-empty", action="store_true", help="保留姓名为空的行（默认不保留）")
    args = ap.parse_args()

    in_path = args.input or find_default_excel()
    if not in_path or not os.path.exists(in_path):
        log("未找到输入 Excel。请把脚本与汇总表放同一目录，或用 -i 指定文件。")
        sys.exit(2)

    out_file = args.out_file or f"按人分Sheet_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    try:
        split_to_sheets(in_path, args.sheet, args.name_col, out_file, args.keep_empty)
    except Exception as e:
        log(f"发生错误：{e}")
        raise


if __name__ == "__main__":
    main()
