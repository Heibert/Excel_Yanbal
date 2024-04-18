"""This script reads an XLSB file using LibreOffice in headless mode and retrieves data from a cell range. """
import os
import subprocess
import time
import uno

# Function to start LibreOffice in headless mode
def start_libreoffice():
    try:
        # Start LibreOffice in headless mode
        subprocess.Popen(['libreoffice', '--headless', '--invisible', '--accept=socket,host=localhost,port=2002;urp;'])
        print("LibreOffice started successfully.")
    except Exception as e:
        print("Error starting LibreOffice:", e)

# Function to retrieve cell range data
def get_cell_range_data(sheet, cell_range):
    cell_range_obj = sheet.getCellRangeByName(cell_range)
    data = []
    for row in cell_range_obj.DataArray:
        row_data = []
        for cell in row:
            row_data.append(cell)
        data.append(row_data)
    return data

# Replace 'file_path' with the path to your XLSX file
file_path = 'ESTADO_DE_CARTERA.xlsb'

if not os.path.exists(file_path):
    print(f'The file {file_path} does not exist.')
else:
    print(f'The file {file_path} exists.')

# Start LibreOffice in headless mode
start_libreoffice()

# Wait for LibreOffice to start (adjust sleep time if needed)
time.sleep(3)

# Connect to LibreOffice
local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

# Load the XLSX file
url = uno.systemPathToFileUrl(os.path.abspath(file_path))
document = desktop.loadComponentFromURL(url, "_blank", 0, ())
if document:
    print(f'File {file_path} loaded successfully.')
else:
    print(f'Error loading file {file_path}.')

# Get a cell range
sheet = document.Sheets.getByName("ESTADO DE CARTERA")

# Function to retrieve data from a column
def get_column_data(sheet, column_index):
    data = []
    max_row = sheet.Rows.Count
    print(f"Max Row: {max_row}")
    row_index = 301
    for i in range(max_row):
        row_index += 1
        cell = sheet.getCellByPosition(column_index, row_index)
        cell_value = cell.getString()
        print(f"Row: {row_index}, Column: {column_index}, Value: {cell_value}")
        if not cell_value:
            break
        data.append(cell_value)
    return data


# Get all the data from the specified column
column_index = 1  # Adjust the column index as needed
column_data = get_column_data(sheet, column_index)
print(column_data)
