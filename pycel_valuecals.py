from monkeypatching import monkeypatch_module_object
from openpyxl.workbook import Workbook
import pycel
from pycel.excelutil import criteria_parser


def _criteria_parser(criteria):
    """
    The original criteria parser function from pycel.excelutil.criteria_parser used in sumifs, countifs, etc.
    does not support matching empty cells. This function monkeypatches the original function to match empty cells
    against an empty string ("").

    TODO: investigate what the Excel default is.
    """
    return criteria_parser(criteria if criteria is not None else "")


def extract_formula_calculations(wb: Workbook):
    """
    Extracts all formula calculations in an openpyxl Workbook and stores them in a dictionary.

    Args:
        wb: An openpyxl Workbook object.

    Returns:
        A dictionary where keys are tuple(sheet_name, cell_coordinate) and values are the evaluated formula values.
    """

    compiler = pycel.ExcelCompiler(excel=wb)
    compiler.recalculate()

    cell_values = {}

    # Ensure custom criteria parser is used
    with monkeypatch_module_object(
        pycel,
        pycel.excelutil.criteria_parser,
        _criteria_parser,
    ):
        for ws in wb:
            ws_name = ws.title
            for row in ws.iter_rows():
                for cell in row:
                    if (
                        cell.value
                        and isinstance(cell.value, str)
                        and cell.value.startswith("=")
                    ):
                        cell_values[(ws_name, cell.coordinate)] = compiler.evaluate(
                            f"{ws_name}!{cell.coordinate}"
                        )

    return cell_values
