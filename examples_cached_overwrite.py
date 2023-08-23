import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from shutil import copyfile

import zipfile
import xml.etree.ElementTree as ET

def set_cached_value(excel_file, sheet_name, cell_address, new_value):
    # Define the XML namespace
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    # 1. Unzip the .xlsx file
    with zipfile.ZipFile(excel_file, 'r') as z:
        # Read the relevant sheet XML into memory
        with z.open(f'xl/worksheets/{sheet_name}.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

    # 2. Find the cell element and update it
    for elem in root.findall('.//main:c', namespaces=ns):  # 'c' is the tag name for cells
        if elem.attrib.get('r') == cell_address:  # 'r' is the attribute that holds the cell address
            value_elem = elem.find('./main:v', namespaces=ns)

            if value_elem is None:
                value_elem = ET.SubElement(elem, 'v')  # Create the 'v' element if it doesn't exist
            value_elem.text = str(new_value)  # Set the cached value

    # 3. Write the updated XML back into the .xlsx file
    with zipfile.ZipFile(excel_file, 'a') as z:
        with z.open(f'xl/worksheets/{sheet_name}.xml', 'w') as f:
            tree.write(f)



# Example usage
copyfile('your_workbook.xlsx', 'your_workbook_cached.xlsx')
set_cached_value('your_workbook_cached.xlsx', 'sheet1', 'A1', 999)
