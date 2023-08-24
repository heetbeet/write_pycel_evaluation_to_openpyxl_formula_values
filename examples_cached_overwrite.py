import openpyxl
from openpyxl import LXML
from openpyxl.cell import _writer as _cell_writer
from openpyxl.worksheet import _writer as _worksheet_writer
from openpyxl.compat import safe_string
from types import SimpleNamespace
from contextlib import contextmanager
from openpyxl.workbook import Workbook
from typing import Dict, Tuple, List, Any, Iterator
from openpyxl import load_workbook
from pathlib import Path
import os
import importlib
from functools import lru_cache


@lru_cache
def _list_modules(package_dir: Path) -> List[str]:
    """
    List all importable Python modules from a package directory.

    Args:
        package_dir: The directory where the package is located.

    Returns:
        A sorted list of importable module strings.
    """

    locations = []
    for path in package_dir.rglob("*.py"):
        parts = path.with_suffix("").relative_to(package_dir.parent).parts

        if path.name == "__init__.py":
            parts = parts[:-1]

        if all([i.isidentifier() for i in parts]):
            locations.append(".".join(parts))

    for i in sorted(locations):
        print(i)
    return sorted(locations)


def _list_monkeypatch_locations(package: Any, function: Any) -> List[Tuple[Any, str]]:
    """
    List all locations where a given function is used in a package.

    Args:
        package: The package in which to search for the function.
        function: The function to look for.

    Returns:
        A list of tuples containing module and attribute name.
    """
    locations = _list_modules(Path(package.__file__).parent)

    f_locations = []
    for location in locations:
        try:
            module = importlib.import_module(location)

        # This is a hack to avoid importing non-modules or modules that cannot be imported
        # because they are reserved for conditional imports like system dependent modules
        # this is dangerous because it can load side effects that are not intended to be loaded
        except Exception:
            continue

        for i, j in module.__dict__.items():
            if j is function:
                f_locations.append((module, i))

    return f_locations


_write_cell_orig = _cell_writer.write_cell


def _write_cell_cache(
    cached_values: Dict[Tuple[str, str], Any],
    xf: Any,
    worksheet: Any,
    cell: Any,
    *args,
    **kwargs
) -> None:
    """
    Custom openpyxl.worksheet._writer.write_cell wrapper function with cached
    values support.

    Args:
        cached_values: An additional argument. This is a dictionary containing
            cached cell values with a mapping of {(worksheet_name, cell_coordinate): value ...}
        xf: Original XML writer object passed by openpyxl.
        worksheet: The original worksheet object passed by openpyxl.
        cell: The original cell to write as passed by openpyxl.
        *args: Any additional args passed by openpyxl.
        **kwargs: Any additional kwargs passed by openpyxl.

    Returns:
        None
    """
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
def _monkey_patch_openpyxl_write_cell(
    cached_values: Dict[Tuple[str, str], Any]
) -> Iterator[None]:
    """
    Context manager to monkeypatch the `write_cell` function in the openpyxl package.

    Args:
        cached_values (Dict[Tuple[str, str], Any]): A dictionary containing cached cell values.

    Yields:
        None: Yields control back to the caller, reverts changes upon exit.
    """
    try:
        # Unfortunately we need to monkeypatch openpyxl everywhere where a `from ... import write_cell` is done
        def _write_cell_closure(*args, **kwargs):
            return _write_cell_cache(cached_values, *args, **kwargs)

        for module, function_name in _list_monkeypatch_locations(
            openpyxl, _write_cell_orig
        ):
            setattr(module, function_name, _write_cell_closure)

        yield
    finally:
        for module, function_name in _list_monkeypatch_locations(
            openpyxl, _write_cell_orig
        ):
            setattr(module, function_name, _write_cell_orig)


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
