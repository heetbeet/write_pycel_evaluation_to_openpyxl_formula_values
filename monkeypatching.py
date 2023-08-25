import openpyxl
from openpyxl import LXML
from openpyxl.cell import _writer as _cell_writer
from openpyxl.compat import safe_string
from openpyxl.xml.functions import whitespace
from types import SimpleNamespace
from contextlib import contextmanager
from openpyxl.workbook import Workbook
from typing import Dict, Tuple, List, Any, Iterator, Union
from openpyxl import load_workbook
from pathlib import Path


def _is_subpath(child_path: Union[Path, str], parent_path: Union[Path, str]) -> bool:
    """
    Check if one path is a subpath of another.

    Args:
        child_path: The child path to check.
        parent_path: The parent path to check against.

    Returns:
        True if child_path is a subpath of parent_path, False otherwise.
    """
    child = Path(child_path).resolve()
    parent = Path(parent_path).resolve()
    try:
        child.relative_to(parent)
        return True
    except ValueError:
        return False


def _module_file_path(module: Any) -> Union[None, Path]:
    """
    Retrieve the file path of a module.

    Args:
        module: The module whose file path is to be retrieved.

    Returns:
        The file path of the module if it exists, otherwise None.
    """
    try:
        return Path(module.__file__).resolve()
    except AttributeError:
        return None


def _find_submodules(module: Any) -> Dict[Path, Any]:
    """
    Recursively find all submodules of a given module.

    Args:
        module: The parent module to search from.

    Returns:
        A dictionary mapping file paths to module objects.

    Raises:
        ValueError: If the module is not loaded from a file.
    """
    cache_dict = {}

    module_path = _module_file_path(module)
    if module_path is None:
        raise ValueError("Module must be a module loaded from a file.")

    module_dir = module_path.parent

    def recurse(module):
        path = _module_file_path(module)
        if (path is None) or (path in cache_dict) or not _is_subpath(path, module_dir):
            return

        cache_dict[path] = module

        for attr_name in dir(module):
            try:
                attr = getattr(module, attr_name)
            except Exception:
                continue

            recurse(attr)

    recurse(module)
    return cache_dict


def _list_monkeypatch_locations(module: Any, obj: Any) -> List[Tuple[Any, str]]:
    """
    List all locations where a given object is used in a module.

    Args:
        module: The module in which to search for the object.
        obj: The object to look for.

    Returns:
        A list of tuples containing module and attribute name.
    """
    locations = _find_submodules(module)

    f_locations = []
    for filename, module in locations.items():
        for attr_name in dir(module):
            if getattr(module, attr_name) is obj:
                f_locations.append((module, attr_name))

    return f_locations


_monkeypatch_module_object_cache = {}


@contextmanager
def monkeypatch_module_object(
    module: Any, obj_original: Any, obj_replacement: Any, cached=False
) -> Iterator[None]:
    """
    Temporarily replace an object within a module and all its submodules.

    This function replaces all instances of `obj_original` in the specified `module`
    and its submodules with `obj_replacement`. After the `with` block, the original object
    is restored. This is useful for temporary monkeypatching for testing or debugging.

    Args:
        module: The module in which to perform the monkeypatching.
        obj_original: The object that will be temporarily replaced.
        obj_replacement: The object that will replace `obj_original`.
        cached: Whether to cache the monkeypatch locations; default is False 
            (recalculation). Warning, this is only correct if  are sure that the 
            monkeypatch locations are still in tact and that the object has not 
            been replaced elsewhere in code.


    Yields:
        None: Yields None while the objects are replaced in the module and its submodules.

    Example:
        >>> with monkeypatch_module_object(math, math.sin, mock_sin):
        ...     assert math.sin(0) == mock_sin(0)
        ...     # other tests or operations that depend on the patched sin function

    Note:
        The function should be used within a `with` statement to ensure that the original
        object is restored even if an error occurs.
    """

    if cached and (path := _module_file_path(module)) is not None:
        if (path, id(obj_original)) not in _monkeypatch_module_object_cache:
            _monkeypatch_module_object_cache[
                (path, id(obj_original))
            ] = _list_monkeypatch_locations(module, obj_original)

        monkeypatch_locations = _monkeypatch_module_object_cache[
            (path, id(obj_original))
        ]
    else:
        monkeypatch_locations = _list_monkeypatch_locations(module, obj_original)

    try:
        for submodule, attr_name in monkeypatch_locations:
            setattr(submodule, attr_name, obj_replacement)
        yield
    finally:
        for submodule, attr_name in monkeypatch_locations:
            setattr(submodule, attr_name, obj_original)
