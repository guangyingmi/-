#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Sheet 拆分工具
自动搜索当前文件夹内的Excel文件，将每个Excel文件中的所有sheet拆分成独立的Excel文件
保持原始sheet的内容、格式和样式完全一致
"""

import os
import sys
import datetime as _dt
import traceback
from pathlib import Path
from typing import List, Optional
from copy import copy

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from tqdm import tqdm


def log(msg: str):
    """日志输出"""
    ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def sanitize_filename(name: str) -> str:
    """清理文件名中的非法字符"""
    import re
    # 替换Windows文件名中的非法字符
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def copy_cell(src, dst):
    """复制单元格的值与样式"""
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def copy_worksheet_with_styles(src_ws, dst_ws):
    """复制工作表内容和样式"""
    # 复制所有单元格
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is not None or cell.has_style:
                dst_cell = dst_ws.cell(row=cell.row, column=cell.column)
                copy_cell(cell, dst_cell)
    
    # 复制列宽
    for col_letter, dimension in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = dimension.width
    
    # 复制行高
    for row_num, dimension in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_num].height = dimension.height
    
    # 复制合并单元格
    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))
    
    # 复制冻结窗格
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes
    
    # 复制自动筛选
    if src_ws.auto_filter:
        dst_ws.auto_filter.ref = src_ws.auto_filter.ref
    
    # 复制打印设置
    dst_ws.page_setup = copy(src_ws.page_setup)
    dst_ws.page_margins = copy(src_ws.page_margins)
    
    # 复制工作表保护
    if src_ws.protection.sheet:
        dst_ws.protection = copy(src_ws.protection)


def find_excel_files(directory: str) -> List[str]:
    """查找目录下的所有Excel文件"""
    excel_files = []
    directory_path = Path(directory)
    
    for file_path in directory_path.iterdir():
        if file_path.is_file():
            # 检查文件扩展名
            if file_path.suffix.lower() in ['.xlsx', '.xlsm', '.xls']:
                # 排除临时文件和输出文件夹
                if not file_path.name.startswith('~$') and '单个Excel合集' not in file_path.name:
                    excel_files.append(str(file_path))
    
    return excel_files


def split_excel_sheets(excel_file: str, output_dir: str) -> int:
    """拆分Excel文件中的所有sheet为独立文件"""
    try:
        log(f"正在处理文件: {os.path.basename(excel_file)}")
        
        # 加载工作簿
        wb = load_workbook(excel_file, data_only=False)
        sheet_count = 0
        
        # 获取所有工作表名称
        sheet_names = wb.sheetnames
        log(f"发现 {len(sheet_names)} 个工作表: {', '.join(sheet_names)}")
        
        # 为每个sheet创建独立的Excel文件
        for sheet_name in tqdm(sheet_names, desc=f"拆分 {os.path.basename(excel_file)}"):
            try:
                # 获取源工作表
                src_ws = wb[sheet_name]
                
                # 创建新的工作簿
                new_wb = Workbook()
                # 删除默认的工作表
                new_wb.remove(new_wb.active)
                
                # 创建新的工作表，使用原始sheet名称
                new_ws = new_wb.create_sheet(title=sheet_name)
                
                # 复制工作表内容和样式
                copy_worksheet_with_styles(src_ws, new_ws)
                
                # 生成输出文件名
                safe_sheet_name = sanitize_filename(sheet_name)
                output_file = os.path.join(output_dir, f"{safe_sheet_name}.xlsx")
                
                # 如果文件已存在，添加序号
                counter = 1
                original_output_file = output_file
                while os.path.exists(output_file):
                    name_without_ext = os.path.splitext(original_output_file)[0]
                    output_file = f"{name_without_ext}_{counter}.xlsx"
                    counter += 1
                
                # 保存文件
                new_wb.save(output_file)
                new_wb.close()
                
                sheet_count += 1
                log(f"已保存: {os.path.basename(output_file)}")
                
            except Exception as e:
                log(f"处理工作表 '{sheet_name}' 时出错: {str(e)}")
                continue
        
        wb.close()
        return sheet_count
        
    except Exception as e:
        log(f"处理文件 '{excel_file}' 时出错: {str(e)}")
        return 0


def main():
    """主函数"""
    try:
        log("Excel Sheet 拆分工具启动")
        
        # 获取当前目录
        current_dir = os.getcwd()
        log(f"当前工作目录: {current_dir}")
        
        # 查找Excel文件
        excel_files = find_excel_files(current_dir)
        
        if not excel_files:
            log("未找到任何Excel文件")
            return
        
        log(f"找到 {len(excel_files)} 个Excel文件:")
        for file in excel_files:
            log(f"  - {os.path.basename(file)}")
        
        # 创建输出目录
        output_dir = os.path.join(current_dir, "单个Excel合集")
        os.makedirs(output_dir, exist_ok=True)
        log(f"输出目录: {output_dir}")
        
        # 处理每个Excel文件
        total_sheets = 0
        successful_files = 0
        
        for excel_file in excel_files:
            try:
                sheet_count = split_excel_sheets(excel_file, output_dir)
                if sheet_count > 0:
                    total_sheets += sheet_count
                    successful_files += 1
                    log(f"文件 '{os.path.basename(excel_file)}' 处理完成，拆分了 {sheet_count} 个工作表")
                else:
                    log(f"文件 '{os.path.basename(excel_file)}' 处理失败")
            except Exception as e:
                log(f"处理文件 '{os.path.basename(excel_file)}' 时发生错误: {str(e)}")
                continue
        
        # 输出总结
        log("\n=== 处理完成 ===")
        log(f"成功处理文件数: {successful_files}/{len(excel_files)}")
        log(f"总共拆分工作表数: {total_sheets}")
        log(f"输出目录: {output_dir}")
        
        if total_sheets > 0:
            log("所有工作表已成功拆分为独立的Excel文件！")
        else:
            log("没有成功拆分任何工作表")
            
    except Exception as e:
        error_msg = f"程序执行出错: {str(e)}\n{traceback.format_exc()}"
        log(error_msg)
        raise


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("\n用户中断程序")
    except Exception as e:
        log(f"程序异常退出: {str(e)}")
        sys.exit(1)