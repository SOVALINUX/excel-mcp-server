# Excel MCP Server Tools

This document provides detailed information about all available tools in the Excel MCP server.

## Workbook Operations

### create_workbook

Creates a new Excel workbook.

```python
create_workbook(filepath: str) -> str
```

- `filepath`: Path where to create workbook
- Returns: Success message with created file path

### create_worksheet

Creates a new worksheet in an existing workbook.

```python
create_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Success message

### get_workbook_metadata

Get metadata about workbook including sheets and ranges.

```python
get_workbook_metadata(filepath: str, include_ranges: bool = False) -> str
```

- `filepath`: Path to Excel file
- `include_ranges`: Whether to include range information
- Returns: String representation of workbook metadata

### read_excel_binary

Read Excel file as base64-encoded binary string for transfer or upload.

```python
read_excel_binary(filepath: str) -> str
```

- `filepath`: Path to Excel file (supports .xlsx, .xlsm, .xlsb, .xls formats)
- Returns: Base64-encoded string of the Excel file binary content

**Use Cases:**
- Upload Excel files to cloud storage (S3, Azure Blob, Google Cloud Storage, etc.)
- Send Excel files through APIs that accept base64-encoded data
- Transfer Excel files through text-based protocols
- Embed Excel files in JSON payloads
- Store Excel files in databases as text fields

**Notes:**
- Returns only the base64-encoded content (not JSON)
- Use `get_workbook_metadata()` to retrieve file metadata (size, sheets, etc.)
- Works with .xlsx, .xlsm, .xlsb (Excel 2007+) and .xls (Excel 97-2003) formats
- The base64 string size will be approximately 33% larger than the original file
- For very large files (>50MB), consider streaming or chunking approaches
- To decode in Python: `base64.b64decode(content)`
- To decode in Node.js: `Buffer.from(content, 'base64')`
- To decode in browser JavaScript: `atob(content)`

### write_excel_binary

Write base64-encoded content to an Excel file.

```python
write_excel_binary(filepath: str, base64_content: str) -> str
```

- `filepath`: Path where to write the Excel file (supports .xlsx, .xlsm, .xlsb, .xls formats)
- `base64_content`: Base64-encoded string of Excel file binary content
- Returns: Success message with file path and size

**Use Cases:**
- Create files from templates stored as base64
- Download files from cloud storage and save locally
- Restore files from database storage
- Write files received from API responses
- Initialize workbooks from pre-existing templates

**Notes:**
- Creates parent directories if they don't exist
- Overwrites existing file at the path
- Works with .xlsx, .xlsm, .xlsb (Excel 2007+) and .xls (Excel 97-2003) formats
- Validates that content is valid base64 and meets minimum size requirements
- Use `read_excel_binary()` to get base64 content from existing files

### delete_file

Delete an Excel file to cleanup and prevent further access.

```python
delete_file(filepath: str) -> str
```

- `filepath`: Path to the Excel file to delete (supports .xlsx, .xlsm, .xlsb, .xls formats)
- Returns: Success message with filepath

**Use Cases:**
- Remove temporary Excel files after processing
- Clean up generated reports or exports
- Delete outdated or obsolete workbooks
- Prevent further access to sensitive files
- Free up disk space by removing unused files

**Notes:**
- File must exist to be deleted
- Requires write permissions on the file and parent directory
- Operation is irreversible - file cannot be recovered
- Works with .xlsx, .xlsm, .xlsb (Excel 2007+) and .xls (Excel 97-2003) formats
- Will raise error if file is currently open or locked by another process

## Data Operations

### write_data_to_excel

Write data to Excel worksheet.

```python
write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[Dict],
    start_cell: str = "A1"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data`: List of dictionaries containing data to write
- `start_cell`: Starting cell (default: "A1")
- Returns: Success message

### read_data_from_excel

Read data from Excel worksheet.

```python
read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    preview_only: bool = False
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default: "A1")
- `end_cell`: Optional ending cell
- `preview_only`: Whether to return only a preview
- Returns: String representation of data

## Formatting Operations

### format_range

Apply formatting to a range of cells.

```python
format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None,
    auto_column_width: bool = False,
    column_width: float = None,
    auto_detect_numeric_columns: bool = False,
    date_format: str = None,
    auto_detect_date_columns: bool = False
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- `bold`: Make text bold
- `italic`: Make text italic
- `underline`: Underline text
- `font_size`: Font size in points
- `font_color`: Font color (hex code)
- `bg_color`: Background color (hex code)
- `border_style`: Border style (thin, medium, thick, double)
- `border_color`: Border color (hex code)
- `number_format`: Excel number format string (e.g., '#,##0.00')
- `alignment`: Text alignment (left, center, right, justify)
- `wrap_text`: Wrap text in cells
- `merge_cells`: Merge the range
- `protection`: Cell protection settings dict
- `conditional_format`: Conditional formatting rules dict
- `auto_column_width`: Auto-adjust column width based on content (approximate, checks longest text including newlines)
- `column_width`: Absolute column width number applied to all columns in range
- `auto_detect_numeric_columns`: Auto-detect and apply number_format to numeric columns. **Converts string numbers to actual numeric values** (e.g., '123' → 123)
- `date_format`: Date format string (e.g., 'yyyy-mm-dd', 'mm/dd/yyyy'). If not specified with auto_detect_date_columns, automatically uses 'yyyy-mm-dd hh:mm:ss' for datetime or 'yyyy-mm-dd' for date-only columns
- `auto_detect_date_columns`: Auto-detect and apply date_format to date columns. **Converts string dates to datetime objects** (e.g., '2025-12-18 08:08:20.000' → datetime). Supports multiple formats including ISO, US, EU, and text formats
- Returns: Success message

**Note on Auto-Detection:**
- When `auto_detect_numeric_columns=True`, the tool scans each column and converts string representations of numbers (like '0', '1', '123') to actual numeric types (int/float)
- **Long Number Protection**: Numbers with more than 15 significant digits are kept as text to prevent data loss. Excel's IEEE-754 double-precision storage has a 15-digit precision limit. Example: Employee IDs like '8760000000000871450' (19 digits) are protected
- When `auto_detect_date_columns=True`, the tool scans each column and converts string representations of dates (like '2025-12-18' or '2025-12-18 08:08:20.000') to datetime objects
- The tool automatically distinguishes between datetime (with time) and date-only columns and applies appropriate formatting
- This ensures Excel properly recognizes and handles the data for sorting, filtering, and calculations
- **Performance Optimization**: Date format is cached per column (detected once, reused for all rows) for 1.8x faster processing

### merge_cells

Merge a range of cells.

```python
merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### unmerge_cells

Unmerge a previously merged range of cells.

```python
unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### get_merged_cells

Get merged cells in a worksheet.

```python
get_merged_cells(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- Returns: String representation of merged cells


## Formula Operations

### apply_formula

Apply Excel formula to cell.

```python
apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to apply
- Returns: Success message

### validate_formula_syntax

Validate Excel formula syntax without applying it.

```python
validate_formula_syntax(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to validate
- Returns: Validation result message

## Chart Operations

### create_chart

Create chart in worksheet.

```python
create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing chart data
- `chart_type`: Type of chart (line, bar, pie, scatter, area)
- `target_cell`: Cell where to place chart
- `title`: Optional chart title
- `x_axis`: Optional X-axis label
- `y_axis`: Optional Y-axis label
- Returns: Success message

## Pivot Table Operations

### create_pivot_table

Create pivot table in worksheet.

```python
create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    target_cell: str,
    rows: List[str],
    values: List[str],
    columns: List[str] = None,
    agg_func: str = "mean"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing source data
- `target_cell`: Cell where to place pivot table
- `rows`: Fields for row labels
- `values`: Fields for values
- `columns`: Optional fields for column labels
- `agg_func`: Aggregation function (sum, count, average, max, min)
- Returns: Success message

## Table Operations

### create_table

Creates a native Excel table from a specified range of data.

```python
create_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: str = None,
    table_style: str = "TableStyleMedium9"
) -> str
```

- `filepath`: Path to the Excel file.
- `sheet_name`: Name of the worksheet.
- `data_range`: The cell range for the table (e.g., "A1:D5").
- `table_name`: Optional unique name for the table.
- `table_style`: Optional visual style for the table.
- Returns: Success message.

## Worksheet Operations

### copy_worksheet

Copy worksheet within workbook.

```python
copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> str
```

- `filepath`: Path to Excel file
- `source_sheet`: Name of sheet to copy
- `target_sheet`: Name for new sheet
- Returns: Success message

### delete_worksheet

Delete worksheet from workbook.

```python
delete_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of sheet to delete
- Returns: Success message

### rename_worksheet

Rename worksheet in workbook.

```python
rename_worksheet(filepath: str, old_name: str, new_name: str) -> str
```

- `filepath`: Path to Excel file
- `old_name`: Current sheet name
- `new_name`: New sheet name
- Returns: Success message

## Range Operations

### copy_range

Copy a range of cells to another location.

```python
copy_range(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Source worksheet name
- `source_start`: Starting cell of source range
- `source_end`: Ending cell of source range
- `target_start`: Starting cell for paste
- `target_sheet`: Optional target worksheet name
- Returns: Success message

### delete_range

Delete a range of cells and shift remaining cells.

```python
delete_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- `shift_direction`: Direction to shift cells ("up" or "left")
- Returns: Success message

### validate_excel_range

Validate if a range exists and is properly formatted.

```python
validate_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Returns: Validation result message

### get_data_validation_info

Get data validation rules and metadata for a worksheet.

```python
get_data_validation_info(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- Returns: JSON string containing all data validation rules with metadata including:
  - Validation type (list, whole, decimal, date, time, textLength)
  - Operator (between, notBetween, equal, greaterThan, lessThan, etc.)
  - Allowed values for list validations (resolved from ranges)
  - Formula constraints for numeric/date validations
  - Cell ranges where validation applies
  - Prompt and error messages

**Note**: The `read_data_from_excel` tool automatically includes validation metadata for individual cells when available.

## Row and Column Operations

### insert_rows

Insert one or more rows starting at the specified row.

```python
insert_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_row`: Row number where to start inserting (1-based)
- `count`: Number of rows to insert (default: 1)
- Returns: Success message

### insert_columns

Insert one or more columns starting at the specified column.

```python
insert_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_col`: Column number where to start inserting (1-based)
- `count`: Number of columns to insert (default: 1)
- Returns: Success message

### delete_sheet_rows

Delete one or more rows starting at the specified row.

```python
delete_sheet_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_row`: Row number where to start deleting (1-based)
- `count`: Number of rows to delete (default: 1)
- Returns: Success message

### delete_sheet_columns

Delete one or more columns starting at the specified column.

```python
delete_sheet_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_col`: Column number where to start deleting (1-based)
- `count`: Number of columns to delete (default: 1)
- Returns: Success message
