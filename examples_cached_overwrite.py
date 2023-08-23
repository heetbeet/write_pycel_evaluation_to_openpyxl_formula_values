from openpyxl import LXML
from openpyxl.cell import _writer as _cell_writer
from openpyxl.worksheet import _writer as _worksheet_writer
from openpyxl.compat import safe_string
from types import SimpleNamespace
from contextlib import contextmanager
from openpyxl.workbook import Workbook
from typing import Dict, Tuple
from openpyxl import load_workbook

_write_cell_orig = _cell_writer.write_cell


def _write_cell(cached_values, xf, worksheet, cell, *args, **kwargs):
    key = (worksheet.title, cell.coordinate)
    if key not in cached_values:
        return _write_cell_orig(xf, worksheet, cell, *args, **kwargs)

    value = cached_values[key]
    if LXML:
        raise NotImplementedError("LXML not supported with cached values")

        ## The following code will result in duplicate "v" elements
        # _write_cell_orig(xf, worksheet, cell, *args, **kwargs)
        # with xf.element("v"):
        #    xf.write(safe_string(value))

    else:
        from xml.etree.ElementTree import SubElement

        ns = SimpleNamespace()

        def write(el):
            ns.el = el

        _write_cell_orig(SimpleNamespace(write=write), worksheet, cell, *args, **kwargs)

        # Add cached value to the "v" element in the cell's XML
        v_elem = ns.el.find("v")
        if v_elem is None:
            SubElement(ns.el, "v")

        v_elem.text = safe_string(value)
        xf.write(ns.el)


@contextmanager
def _monkey_patch_openpyxl_write_cell(cached_values):
    try:
        # Unfortunately we need to monkeypatch openpyxl everywhere where a `from ... import write_cell` is done
        _cell_writer.write_cell = lambda *args, **kwargs: _write_cell(
            cached_values, *args, **kwargs
        )
        _worksheet_writer.write_cell = lambda *args, **kwargs: _write_cell(
            cached_values, *args, **kwargs
        )
        yield
    finally:
        _cell_writer.write_cell = _write_cell_orig
        _worksheet_writer.write_cell = _write_cell_orig


def save_workbook_with_cache(
    workbook: Workbook, filename: str, cached_values: Dict[Tuple[str, str], int]
) -> None:
    """
    Save an Openpyxl Workbook with certain cell values overridden by cached values.
    This allows you to prepopulate certain cells with values without having to
    calculate the workbook in an external application.

    This function saves a workbook after applying cached values to specific cells.
    The cached values are specified in a dictionary where the keys are tuples
    consisting of the worksheet name and cell coordinate, and the values are the
    cached values to apply.

    Args:
        workbook (Workbook): The Openpyxl Workbook object to save.
        filename (str): The name of the file to save the workbook as.
        cached_values (Dict[Tuple[str, str], int]): Dictionary containing the cached values.
            The keys are tuples where the first element is the worksheet name and the
            second element is the cell coordinate (e.g., ('Sheet1', 'A1')). The values
            are the cached values to set for those cells.

    Returns:
        None
    """

    with _monkey_patch_openpyxl_write_cell(cached_values):
        return workbook.save(filename)


if __name__ == "__main__":
    # Move this to tests or something

    cached_values = {}
    cached_values[("Sheet1", "A1")] = 9999

    workbook = load_workbook("your_workbook.xlsx")
    workbook.calculation.calcMode = "manual"
    save_workbook_with_cache(workbook, "your_workbook_cached.xlsx", cached_values)

    # openpyxl read cached workbook data only and see what A1 is
    workbook = load_workbook(
        "your_workbook_cached.xlsx", read_only=True, data_only=True
    )
    print(workbook["Sheet1"]["A1"].value)
