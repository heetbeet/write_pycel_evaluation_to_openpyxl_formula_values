from pycel_valuecals import extract_formula_calculations
from openpyxl_replacetable import replace_table
from openpyxl_valuecache import save_workbook_with_cache

import pandas as pd
import openpyxl

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
cell_values = extract_formula_calculations(wb)

# set wb to not calculate formulas
wb.calculation.calcMode = "manual"
save_workbook_with_cache(wb, "va-template-cached.xlsx", cell_values)
