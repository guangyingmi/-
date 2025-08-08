
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 汇总表「按人拆分」工具
---------------------------------
用途：
    将一个汇总 Excel（含所有人的记录）按“姓名/销售员”等列拆分为多个 Excel 文件，
    每个人一个文件，文件名即为该人姓名。自动将“xxx 汇总”行一并归入对应人员文件。

使用方法：
    方式一：双击同目录下的 EXE（由本脚本打包生成），自动扫描当前目录的 xlsx 文件并拆分。
    方式二：命令行运行：
        python split_by_person.py -i "汇总.xlsx" -c 销售员

参数：
    -i, --input       输入 Excel 文件路径（默认自动扫描当前目录 *.xlsx 的第一个）
    -s, --sheet       工作表名或索引（默认 0，即第一张）
    -c, --name-col    姓名列名（默认自动从列名中匹配：销售员/姓名/员工/人员/负责人/Name）
    -o, --out-dir     输出目录（默认：按人拆分_YYYYMMDD_HHMMSS）
    --keep-empty      保留姓名为空的行（默认不保留）

依赖：pandas、openpyxl
打包：pyinstaller --onefile --name Excel按人拆分 split_by_person.py

作者：ChatGPT 生成
版本：1.0.0
"""
import argparse
import sys
import os
import re
import datetime as _dt
from typing import Optional, List, Union
import pandas as pd

DEFAULT_NAME_KEYS = ["销售员","姓名","员工","人员","负责人","Name","name"]

def log(msg: str):
    ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def find_default_excel() -> Optional[str]:
    # 选择当前目录下第一个正常的 xlsx
    for name in os.listdir("."):
        if name.lower().endswith(".xlsx") and not name.startswith("~$"):
            # 排除可能的结果输出文件
            if "按人拆分" in name:
                continue
            return name
    return None

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Excel 汇总表按人拆分工具")
    p.add_argument("-i","--input", help="输入 Excel 路径（默认自动扫描）")
    p.add_argument("-s","--sheet", help="工作表名或索引（默认 0）", default="0")
    p.add_argument("-c","--name-col", help="姓名列名（默认自动匹配）")
    p.add_argument("-o","--out-dir", help="输出目录")
    p.add_argument("--keep-empty", action="store_true", help="保留姓名为空的行")
    return p.parse_args()

def load_excel(path: str, sheet: str) -> pd.DataFrame:
    # sheet 可以是名称或字符串数字索引
    try:
        if sheet.isdigit():
            df = pd.read_excel(path, sheet_name=int(sheet))
        else:
            df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        # 回退：默认第一张
        df = pd.read_excel(path)
    # 统一清理列名
    df.columns = [str(c).strip() for c in df.columns]
    return df

def detect_name_col(columns: List[str], manual: Optional[str]=None) -> str:
    if manual and manual in columns:
        return manual
    # 优先完整匹配
    for key in DEFAULT_NAME_KEYS:
        if key in columns:
            return key
    # 次选包含匹配
    for c in columns:
        if any(key in c for key in DEFAULT_NAME_KEYS):
            return c
    # 否则用第一列兜底
    return columns[0]

def base_name(s: Union[str, float, int]) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s*汇总$", "", s)  # 去掉末尾的“汇总”
    return s.strip()

def sanitize_filename(name: str) -> str:
    # Windows 不允许的字符替换为下划线
    return re.sub(r'[\\/:*?"<>|]', "_", name)

def split_by_person(df: pd.DataFrame, name_col: str, keep_empty: bool=False) -> int:
    df["_base_name"] = df[name_col].apply(base_name)
    if not keep_empty:
        df = df[df["_base_name"].astype(bool)]
    # 保持原有顺序的 groupby：使用 groupby(sort=False)
    count = 0
    for person, sub in df.groupby("_base_name", sort=False):
        out = sub.drop(columns=["_base_name"])
        fname = sanitize_filename(person) or "未命名"
        yield fname, out
        count += 1
    return count

def main():
    args = parse_args()
    in_path = args.input or find_default_excel()
    if not in_path or not os.path.exists(in_path):
        log("未找到输入 Excel。请将 EXE 与汇总表放在同一目录，或通过 -i 指定文件。")
        sys.exit(2)

    log(f"输入文件：{in_path}")
    df = load_excel(in_path, args.sheet)
    if df.empty:
        log("工作表为空。程序结束。")
        sys.exit(0)

    name_col = detect_name_col(df.columns.tolist(), args.name_col)
    if name_col not in df.columns:
        log(f"未找到姓名列：{name_col}")
        sys.exit(3)
    log(f"使用姓名列：{name_col}")

    out_dir = args.out_dir or f"按人拆分_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    os.makedirs(out_dir, exist_ok=True)
    log(f"输出目录：{out_dir}")

    total = 0
    from pandas import ExcelWriter
    for fname, sub in split_by_person(df, name_col, args.keep_empty):
        path = os.path.join(out_dir, f"{fname}.xlsx")
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            sub.to_excel(writer, index=False)
        total += 1
        log(f"已生成：{path}（{len(sub)} 行）")

    log(f"完成！共生成 {total} 个文件。")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"发生错误：{e}")
        raise
