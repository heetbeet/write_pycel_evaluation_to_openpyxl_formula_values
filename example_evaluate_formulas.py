import openpyxl
from pycel import ExcelCompiler
from monkeypatching import monkeypatch_module_object
from openpyxl_valuecache import save_workbook_with_cache
from openpyxl.utils import coordinate_to_tuple
import pandas as pd
import pycel

from pycel.excelutil import criteria_parser


def _criteria_parser(criteria):
    if criteria is None:
        criteria = ""
    return criteria_parser(criteria)


def replace_table(wb, tablename, df):
    # Initialize variables
    table = None
    ws = None

    # Loop through all worksheets to find the table
    for sheet in wb:
        for tbl in sheet._tables.values():
            if tbl.name == tablename:
                table = tbl
                ws = sheet
                break
        if table:
            break

    # If table is not found, throw an exception
    if table is None:
        raise Exception("Table not found.")

    # Parse table.ref string to get the start and end cells
    start_cell, end_cell = table.ref.split(":")
    start_row, start_col = coordinate_to_tuple(start_cell)
    end_row, end_col = coordinate_to_tuple(end_cell)

    # Clear existing data
    for row in ws.iter_rows(
        min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col
    ):
        for cell in row:
            cell.value = None

    # Update table headers and data
    new_headers = list(df.columns)
    new_data = df.values.tolist()
    all_data = [new_headers] + new_data

    # Write all data to the worksheet
    for i, row_data in enumerate(all_data, start=start_row):
        for j, value in enumerate(row_data, start=start_col):
            ws.cell(row=i, column=j, value=value)

    # Resize table to fit new data
    end_row = start_row + len(all_data) - 1
    end_col = start_col + len(new_headers) - 1
    new_range = (
        ws.cell(row=start_row, column=start_col).coordinate
        + ":"
        + ws.cell(row=end_row, column=end_col).coordinate
    )
    table.ref = new_range


# Load the Excel workbook
wb = openpyxl.load_workbook("va-template.xlsx")

# Create DataFrame
data = {
    "Virtual Actuary name": ["A", "B", "A", "C", "D", "B", "A", "C", "D", "A"],
    "Virtual Actuary ID": [
        "VA-001",
        "VA-002",
        "VA-001",
        "VA-003",
        "VA-004",
        "VA-002",
        "VA-001",
        "VA-003",
        "VA-004",
        "VA-001",
    ],
    "client": [
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
        "Sanlam",
    ],
    "project": [
        "SEM_GI_Support",
        "SEM_GI_TM1Testing",
        "SEM_GI_Support",
        "SEM_GI_ActuarialBAU",
        "SEM_GI_Support",
        "SEM_GI_TM1Testing",
        "SEM_GI_Support",
        "SEM_GI_TM1Testing",
        "SEM_GI_ActuarialBAU",
        "SEM_GI_ActuarialBAU",
    ],
    "tags": [None, None, None, None, None, None, None, None, None, None],
    "week_starting": [
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
        "2023-06-26",
    ],
    "description": [
        "May run",
        "May run",
        "May run",
        "May run",
        "May run",
        "May run",
        "May run",
        "May run",
        "May run",
        "May run",
    ],
    "date": [
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
        "2023-06-28",
    ],
    "duration": [1.32, 2.51, 0.80, 2.89, 1.5, 0.7, 1.2, 1.0, 1.3, 0.9],
    "retrieval_date": [
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
        "2023-06-30",
    ],
}

df = pd.DataFrame(data)

# Replace table
replace_table(wb, "Timesheet", df)

with monkeypatch_module_object(pycel, criteria_parser, _criteria_parser):
    compiler = ExcelCompiler(excel=wb)

    compiler.recalculate()

    # Create an empty dictionary to store cell coordinates and their evaluated values
    cell_values = {}

    # Loop through the worksheets in the workbook
    for ws in wb:
        ws_name = ws.title

        # Loop through all the cells in the worksheet
        for row in ws.iter_rows():
            for cell in row:
                # Check if the cell has a formula
                if (
                    cell.value
                    and isinstance(cell.value, str)
                    and cell.value.startswith("=")
                ):
                    formula = cell.value[1:]  # Remove the "="
                    coord = f"{ws_name}!{cell.coordinate}"

                    cell_values[ws_name, cell.coordinate] = compiler.evaluate(coord)


# set wb to not calculate formulas
wb.calculation.calcMode = "manual"
save_workbook_with_cache(wb, "va-template-cached.xlsx", cell_values)
