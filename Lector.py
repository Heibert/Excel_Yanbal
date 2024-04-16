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

# Replace 'file_path' with the path to your XLSB file
file_path = 'ESTADO_DE_CARTERA.xlsb'

if not os.path.exists(file_path):
    print(f'The file {file_path} does not exist.')
else:
    print(f'The file {file_path} exists.')

# Start LibreOffice in headless mode
start_libreoffice()

# Wait for LibreOffice to start (adjust sleep time if needed)
time.sleep(5)

try:
    # Open the XLSB file to process it
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    context = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
    )
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    url = uno.systemPathToFileUrl(os.path.abspath(file_path))
    doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
    print("File opened successfully.")

    # Retrieve the data from the XLSB file
    sheet = doc.Sheets.getByIndex(0)  # Adjust the index as per your sheet
    cell_range = sheet.getCellRangeByName("C6")

    # Create export options for PDF
    export_props = (
        uno.createUnoStruct("com.sun.star.beans.PropertyValue"),
    )
    export_props[0].Name = "FilterName"
    export_props[0].Value = "calc_pdf_Export"

    # Export the selected cell range to a PDF file
    output_file_path = 'output.pdf'
    output_url = uno.systemPathToFileUrl(os.path.abspath(output_file_path))
    doc.storeToURL(output_url, export_props)
    print(f'File exported to {output_file_path}.')

    # Close the XLSB file
    doc.close(True)
    print("File closed successfully.")
except Exception as e:
    print("Error processing the file:", e)
