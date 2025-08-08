#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 汇总表「按人拆分」（保留样式，无序号前缀）
---------------------------------------------
- 保留样式：表头背景色、单元格边框/字体/对齐、列宽、行高、冻结窗格、自动筛选
- 自动识别姓名列（销售员/姓名/员工/人员/负责人/Name），也可用参数指定
- 自动将“某某 汇总”行并入对应人员文件

依赖：openpyxl>=3.1
打包示例（PyInstaller）：
    pyinstaller --onefile --name "Excel按人拆分" split_by_person_styles.py
"""
import argparse
import os
import re
import sys
from collections import OrderedDict
from copy import copy
import datetime as _dt
from typing import Optional, List, Dict

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

DEFAULT_NAME_KEYS = ["销售员","姓名","员工","人员","负责人","Name","name"]


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


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Excel 汇总表按人拆分（保留样式）")
    p.add_argument("-i","--input", help="输入 Excel 路径（默认自动扫描）")
    p.add_argument("-s","--sheet", help="表名（精确匹配）或索引（0 基）", default=None)
    p.add_argument("-c","--name-col", help="姓名列名（默认自动识别）")
    p.add_argument("-o","--out-dir", help="输出目录")
    p.add_argument("--keep-empty", action="store_true", help="保留姓名为空的行（默认不保留）")
    return p.parse_args()


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
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"\s*汇总$", "", s)  # 去掉“ 汇总”后缀
    return s.strip()


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def copy_cell(src, dst):
    """复制单元格的值与样式（字体、对齐、边框、填充、数字格式、保护）。"""
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)              # 字体
        dst.border = copy(src.border)          # 边框
        dst.fill = copy(src.fill)              # 背景填充
        dst.number_format = src.number_format  # 数字格式
        dst.protection = copy(src.protection)  # 保护
        dst.alignment = copy(src.alignment)    # 对齐方式（水平/垂直/换行等）


def copy_header_and_dimensions(src_ws, dst_ws, header_row=1):
    """复制列宽、表头样式、行高、冻结窗格。"""
    # 列宽
    for col_letter, dim in src_ws.column_dimensions.items():
        nd = dst_ws.column_dimensions[col_letter]
        nd.width = dim.width
        nd.hidden = dim.hidden
        nd.bestFit = dim.bestFit

    # 表头（源表 header_row -> 目标第 1 行）
    for col in range(1, src_ws.max_column + 1):
        sc = src_ws.cell(row=header_row, column=col)
        dc = dst_ws.cell(row=1, column=col)
        copy_cell(sc, dc)

    # 行高
    try:
        dst_ws.row_dimensions[1].height = src_ws.row_dimensions[header_row].height
    except Exception:
        pass

    # 冻结窗格（例如 A2）
    dst_ws.freeze_panes = src_ws.freeze_panes


def write_row_from_src(src_ws, dst_ws, src_row_idx, dst_row_idx):
    """按列复制一整行（值+样式）。"""
    for col in range(1, src_ws.max_column + 1):
        sc = src_ws.cell(row=src_row_idx, column=col)
        dc = dst_ws.cell(row=dst_row_idx, column=col)
        copy_cell(sc, dc)


def run(input_path: Optional[str], sheet_sel, name_col_manual: Optional[str],
        out_dir: Optional[str], keep_empty: bool):
    in_path = input_path or find_default_excel()
    if not in_path or not os.path.exists(in_path):
        log("未找到输入 Excel。请将 EXE 与汇总表放同一目录，或用 -i 指定路径。")
        sys.exit(2)

    log(f"输入文件：{in_path}")
    # data_only=True：读取公式显示值；样式依然可用
    wb = load_workbook(in_path, data_only=True)
    ws = detect_sheet(wb, sheet_sel)
    log(f"工作表：{ws.title}")

    # 表头（第 1 行）
    header = [str(c.value).strip() if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    if not header or all(not h for h in header):
        log("无法读取表头（第 1 行为空）。")
        sys.exit(3)

    name_col = detect_name_col(header, name_col_manual)
    try:
        name_col_idx = header.index(name_col) + 1  # 1-based
    except ValueError:
        log(f"未找到姓名列：{name_col}")
        sys.exit(4)
    log(f"使用姓名列：{name_col}（第 {name_col_idx} 列）")

    out_dir = out_dir or f"按人拆分_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(out_dir, exist_ok=True)
    log(f"输出目录：{out_dir}")

    # 人员 -> （workbook, worksheet），保持插入顺序（即首次出现顺序）
    books: Dict[str, Dict] = OrderedDict()
    header_row_idx = 1

    # 遍历数据行（从第 2 行开始）
    for r in range(2, ws.max_row + 1):
        person_raw = ws.cell(row=r, column=name_col_idx).value
        person = base_name(person_raw)
        if not person and not keep_empty:
            continue

        if person not in books:
            new_wb = Workbook()
            new_ws = new_wb.active
            new_ws.title = person or "未命名"
            copy_header_and_dimensions(ws, new_ws, header_row=header_row_idx)
            books[person] = {"wb": new_wb, "ws": new_ws}

        dst_ws = books[person]["ws"]
        dst_next_row = dst_ws.max_row + 1
        write_row_from_src(ws, dst_ws, r, dst_next_row)

    # 设置筛选范围并保存
    total = 0
    for person, info in books.items():
        wb2, ws2 = info["wb"], info["ws"]
        last_col_letter = get_column_letter(ws2.max_column)
        ws2.auto_filter.ref = f"A1:{last_col_letter}{ws2.max_row}"

        safe = sanitize_filename(person) or "未命名"
        out_path = os.path.join(out_dir, f"{safe}.xlsx")
        wb2.save(out_path)
        total += 1
        log(f"已生成：{out_path}（{ws2.max_row-1} 条记录）")

    log(f"完成！共生成 {total} 个文件。")


def main():
    args = parse_args()
    try:
        run(args.input, args.sheet, args.name_col, args.out_dir, args.keep_empty)
    except Exception as e:
        log(f"发生错误：{e}")
        raise


if __name__ == "__main__":
    main()
