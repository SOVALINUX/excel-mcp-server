import re
from typing import TYPE_CHECKING

from openpyxl.utils import column_index_from_string

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

def parse_cell_range(
    cell_ref: str,
    end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """Parse Excel cell reference into row and column indices."""
    if end_ref:
        start_cell = cell_ref
        end_cell = end_ref
    else:
        start_cell = cell_ref
        end_cell = None

    match = re.match(r"([A-Z]+)([0-9]+)", start_cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {start_cell}")
    col_str, row_str = match.groups()
    start_row = int(row_str)
    start_col = column_index_from_string(col_str)

    if end_cell:
        match = re.match(r"([A-Z]+)([0-9]+)", end_cell.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {end_cell}")
        col_str, row_str = match.groups()
        end_row = int(row_str)
        end_col = column_index_from_string(col_str)
    else:
        end_row = None
        end_col = None

    return start_row, start_col, end_row, end_col


def get_actual_data_range(
    sheet: "Worksheet",
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    max_empty_rows: int = 10
) -> tuple[int, int]:
    """Find the actual data range within a specified range.
    
    Scans the worksheet to find the last row with data, stopping early
    if it encounters consecutive empty rows beyond the data.
    
    Args:
        sheet: The worksheet to scan
        start_row: Starting row index
        start_col: Starting column index
        end_row: Ending row index
        end_col: Ending column index
        max_empty_rows: Stop scanning after this many consecutive empty rows
        
    Returns:
        Tuple of (max_data_row, max_data_col) - the actual extent of data
        
    Example:
        If data exists in rows 1-5 but range is A1:Z10000,
        returns (5, 26) to avoid scanning 9995 empty rows
    """
    max_data_row = start_row
    max_data_col = start_col
    empty_row_count = 0
    
    for row in range(start_row, end_row + 1):
        has_content = False
        row_max_col = start_col
        
        for col in range(start_col, end_col + 1):
            cell_value = sheet.cell(row=row, column=col).value
            if cell_value is not None and cell_value != '':
                has_content = True
                max_data_row = row
                row_max_col = col
                max_data_col = max(max_data_col, col)
        
        if has_content:
            empty_row_count = 0
        else:
            empty_row_count += 1
            # Stop scanning after max_empty_rows consecutive empty rows
            if empty_row_count >= max_empty_rows and row > max_data_row:
                break
    
    return max_data_row, max_data_col

def validate_cell_reference(cell_ref: str) -> bool:
    """Validate Excel cell reference format (e.g., 'A1', 'BC123')"""
    if not cell_ref:
        return False

    # Split into column and row parts
    col = row = ""
    for c in cell_ref:
        if c.isalpha():
            if row:  # Letters after numbers not allowed
                return False
            col += c
        elif c.isdigit():
            row += c
        else:
            return False

    return bool(col and row) 