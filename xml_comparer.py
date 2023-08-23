import zipfile
import os
import xml.etree.ElementTree as ET
import tempfile
from pathlib import Path
from contextlib import suppress

def dump_sheet1(filepath):
    # Ensure the dump directory exists
    if not os.path.exists("dump"):
        os.makedirs("dump")

    # Open the Excel zip archive
    with tempfile.TemporaryDirectory() as tdir:
        with zipfile.ZipFile(filepath, "r") as z:
            # Extract 'xl/worksheets/sheet1.xml' to the 'dump' directory
            z.extract("xl/worksheets/sheet1.xml", tdir)

        with suppress(FileNotFoundError):
            os.remove("dump/sheet1.xml")

        list(Path(tdir).rglob("sheet1.xml"))[0].rename("dump/sheet1.xml")


def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def pretty_print_xml(xml_file):
    # Parse the XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Indent the XML elements
    indent(root)

    with suppress(FileNotFoundError):
        os.remove(xml_file)
    tree.write(xml_file, encoding="utf-8", xml_declaration=True)


#dump_sheet1("your_workbook_cached.xlsx")

# Example usage
for i in list(Path("dump").rglob("*.xml")) + list(Path("dump").rglob("*.rels")):
    pretty_print_xml(i)

