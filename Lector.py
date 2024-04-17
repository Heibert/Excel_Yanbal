import os
import subprocess
import time
from PIL import ImageGrab
from fpdf import FPDF
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

# Function to export cell range to PDF with screenshot-like appearance
def export_to_pdf_screenshot(data, left, top, right, bottom):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    screenshot_image = "screenshot.png"
    data_image = ImageGrab.grab(bbox=(left, top, right, bottom))  # Capture screenshot of cell range
    data_image.save(screenshot_image)
    pdf.image(screenshot_image, x=10, y=10, w=180)
    pdf_output = "output_screenshot.pdf"
    pdf.output(pdf_output)
    print(f"PDF with screenshot-like appearance created successfully: {pdf_output}")

# Replace 'file_path' with the path to your XLSX file
file_path = 'ESTADO_DE_CARTERA.xlsb'

if not os.path.exists(file_path):
    print(f'The file {file_path} does not exist.')
else:
    print(f'The file {file_path} exists.')

# Start LibreOffice in headless mode
start_libreoffice()

# Wait for LibreOffice to start (adjust sleep time if needed)
time.sleep(5)

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
sheet = document.Sheets.getByIndex(2)
# Modify the cell range as needed
cell_range = "B6:H17"

# Get position of the cell range
cell_range_obj = sheet.getCellRangeByName(cell_range)
left, top, right, bottom = cell_range_obj.RangeAddress.StartColumn * 64, cell_range_obj.RangeAddress.StartRow * 20, cell_range_obj.RangeAddress.EndColumn * 64, cell_range_obj.RangeAddress.EndRow * 20

# Get data from cell range
data = get_cell_range_data(sheet, cell_range)

# Recalculate the formulas
document.calculateAll()

# Export cell range to PDF with screenshot-like appearance
export_to_pdf_screenshot(data, left, top, right, bottom)

# Close the document
document.close(True)

# Close LibreOffice
subprocess.Popen(['pkill', 'soffice.bin'])
