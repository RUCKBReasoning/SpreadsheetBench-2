#!/usr/bin/env python3
"""
view_xlsx - View contents of Excel (.xlsx) files
Supports listing sheets and viewing sheet contents with optional row ranges
"""
import sys
import argparse
import pathlib
from typing import Optional


def check_openpyxl():
    """Check if openpyxl is available, provide installation instructions if not"""
    try:
        import openpyxl
        return True
    except ImportError:
        print("Error: openpyxl library is required but not installed.")
        print("Please install it using: pip install openpyxl")
        return False


def get_data_range(sheet):
    """
    Determine the actual data range in a sheet (min row to max row with data)
    Returns tuple (min_row, max_row, min_col, max_col)
    """
    if sheet.max_row is None or sheet.max_column is None:
        return None, None, None, None
    
    # Find the actual data boundaries
    min_row = None
    max_row = None
    min_col = None
    max_col = None
    
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                if min_row is None or cell.row < min_row:
                    min_row = cell.row
                if max_row is None or cell.row > max_row:
                    max_row = cell.row
                if min_col is None or cell.column < min_col:
                    min_col = cell.column
                if max_col is None or cell.column > max_col:
                    max_col = cell.column
    
    return min_row, max_row, min_col, max_col


def list_sheets(workbook):
    """List visible sheet names in the workbook (exclude hidden and veryHidden sheets)"""
    hidden_states = ('hidden', 'veryHidden')
    visible_sheets = [name for name in workbook.sheetnames if workbook[name].sheet_state not in hidden_states]
    print(f"Sheets: {visible_sheets}")


def format_cell_value(value):
    """Format cell value for display"""
    if value is None:
        return ""
    return str(value)


def view_sheet_content(sheet, start_row: Optional[int] = None, end_row: Optional[int] = None):
    """
    View content of a specific sheet
    If start_row and end_row are None, automatically detect data range
    Outputs content in chunks of 100 rows to avoid truncation
    """
    from openpyxl.utils import get_column_letter
    
    # Get actual data range
    min_row, max_row, min_col, max_col = get_data_range(sheet)
    
    if min_row is None:
        print(f"Sheet '{sheet.title}' is empty (no data found)")
        return
    
    # Use provided range or default to detected range
    actual_start = start_row if start_row is not None else min_row
    actual_end = end_row if end_row is not None else max_row
    
    # Validate range
    if actual_start < 1:
        actual_start = 1
    if actual_end < actual_start:
        print(f"Error: end_row ({actual_end}) must be >= start_row ({actual_start})")
        return
    
    # Get column range
    start_col = min_col if min_col else 1
    end_col = max_col if max_col else sheet.max_column
    
    # Convert to Excel-style range notation
    start_col_letter = get_column_letter(start_col)
    end_col_letter = get_column_letter(end_col)
    range_notation = f"'{sheet.title}'!{start_col_letter}{actual_start}:{end_col_letter}{actual_end}"
    
    print(f"\nSheet: {sheet.title}")
    print(f"Data range: {range_notation}")
    print(f"Rows: {actual_start} to {actual_end} (total: {actual_end - actual_start + 1})")
    print(f"Columns: {start_col_letter} to {end_col_letter}")
    print()
    
    # Process in chunks of 100 rows
    CHUNK_SIZE = 100
    current_row = actual_start
    
    while current_row <= actual_end:
        chunk_end = min(current_row + CHUNK_SIZE - 1, actual_end)
        
        if actual_end - actual_start + 1 > CHUNK_SIZE:
            print(f"[Rows {current_row} to {chunk_end}]")
        
        # Print data rows - display as Python list format
        for row_idx in range(current_row, chunk_end + 1):
            row_values = []
            for col_idx in range(start_col, end_col + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                # 获取原始值（公式或数据）
                value = cell.value
                row_values.append(value)
            # 以 Python 列表格式输出，完整显示不截断
            print(f"Row {row_idx}: {row_values}")
        
        current_row = chunk_end + 1
        
        # Add separator between chunks
        if current_row <= actual_end:
            print()


def main():
    parser = argparse.ArgumentParser(
        description="View contents of Excel (.xlsx) files",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("file_path", help="Path to the .xlsx file")
    # 支持两种方式：带 --flag 的形式和纯位置参数形式
    parser.add_argument("--mode", choices=["list", "content"], dest="mode_opt",
                        help="Mode: 'list' to list sheets, 'content' to view content (default: content)")
    parser.add_argument("--sheet", dest="sheet_opt", help="Sheet name to view (default: first visible sheet)")
    parser.add_argument("--start_row", type=int, dest="start_row_opt", help="Starting row number (1-based)")
    parser.add_argument("--end_row", type=int, dest="end_row_opt", help="Ending row number (1-based)")
    # 位置参数形式：view_xlsx file_path [mode] [sheets] [start_row] [end_row]
    parser.add_argument("mode", nargs="?", help="Mode (optional positional: 'list' or 'content')")
    parser.add_argument("sheet", nargs="?", help="Sheet name to view (optional positional)")
    parser.add_argument("start_row", nargs="?", type=int, help="Starting row number (optional positional)")
    parser.add_argument("end_row", nargs="?", type=int, help="Ending row number (optional positional)")
    
    args = parser.parse_args()
    
    # 统一参数来源
    raw_mode = args.mode_opt or args.mode
    sheet = args.sheet_opt or args.sheet
    start_row = args.start_row_opt if args.start_row_opt is not None else args.start_row
    end_row = args.end_row_opt if args.end_row_opt is not None else args.end_row
    
    # 处理 mode 与 sheet 的歧义：如果位置参数 mode 不是合法的模式，就把它当作 sheet
    if raw_mode in (None, "list", "content"):
        mode = raw_mode or "content"
    else:
        mode = "content"
        if sheet is None:
            sheet = raw_mode
    
    # # 检测可能的参数解析问题：sheet 名称以数字结尾但没有 start_row
    # # 这可能是因为 sheet 名称末尾的空格和后续参数被错误合并
    # import re
    # if sheets and start_row is None and re.search(r'\s+\d+$', sheets):
    #     print(f"Warning: Sheet name '{sheets}' ends with a number after whitespace.")
    #     print("This might indicate incorrect argument parsing.")
    #     print("If the sheet name has trailing spaces, make sure they are properly quoted.")
    #     print(f"Example: view_xlsx file.xlsx content '{sheets.rsplit(None, 1)[0]} ' {sheets.rsplit(None, 1)[1]}")
    #     print()
    
    # Check if openpyxl is available
    if not check_openpyxl():
        sys.exit(1)
    
    # Import openpyxl after checking availability
    import openpyxl
    
    # Validate file path
    file_path = pathlib.Path(args.file_path)
    if not file_path.exists():
        print(f"Error: File '{file_path}' does not exist")
        sys.exit(1)
    
    if not file_path.is_file():
        print(f"Error: '{file_path}' is not a file")
        sys.exit(1)
    
    if file_path.suffix.lower() not in ['.xlsx', '.xlsm']:
        print(f"Error: File must be an Excel file (.xlsx or .xlsm), got: {file_path.suffix}")
        sys.exit(1)
    
    try:
        # Load workbook - 不使用 data_only，显示原始公式
        print(f"Loading workbook: {file_path}")
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        
        if mode == "list":
            # List all sheets
            list_sheets(workbook)
        else:
            # View content mode - 只支持查看单个 sheet
            # 获取可见 sheet 列表（排除隐藏的 sheet）
            hidden_states = ('hidden', 'veryHidden')
            visible_sheets = [name for name in workbook.sheetnames if workbook[name].sheet_state not in hidden_states]
            
            if sheet:
                # 用户指定了 sheet 名称，进行智能匹配
                sheet_name = sheet
                matched_name = None
                if sheet_name in workbook.sheetnames:
                    matched_name = sheet_name
                else:
                    # 尝试 strip 后精确匹配
                    stripped_name = sheet_name.strip()
                    if stripped_name in workbook.sheetnames:
                        matched_name = stripped_name
                    else:
                        # 尝试找到 strip 后相等的 sheet（处理 sheet name 本身有首尾空格的情况）
                        for actual_name in workbook.sheetnames:
                            if actual_name.strip() == stripped_name:
                                matched_name = actual_name
                                break
                
                if matched_name is None:
                    print(f"\nError: Sheet '{sheet_name}' not found in workbook")
                    print(f"Available sheets: {visible_sheets}")
                    sys.exit(1)
                
                target_sheet = workbook[matched_name]
            else:
                # 未指定 sheet，使用第一个可见 sheet
                if not visible_sheets:
                    print("Error: No visible sheets in workbook")
                    sys.exit(1)
                target_sheet = workbook[visible_sheets[0]]
            
            view_sheet_content(target_sheet, start_row, end_row)
        
        workbook.close()
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
