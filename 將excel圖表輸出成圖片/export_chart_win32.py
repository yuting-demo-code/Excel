import win32com.client as win32
import os
from time import sleep


def export_charts_as_png(excel_file_path, output_folder):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want to see the Excel application
    excel.DisplayAlerts = False

    workbook = excel.Workbooks.Open(excel_file_path)
    worksheets = workbook.Worksheets

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    for worksheet in worksheets:
        for i, chart in enumerate(worksheet.ChartObjects()):
            # Export each chart as .png
            chart.CopyPicture()
            chart.Chart.Export(os.path.join(os.getcwd(), output_folder, f"chart{i}.png"))
        break

    # sleep(1)
    workbook.Close(SaveChanges=False, Filename=excel_file_path)
    # workbook.Close()
    excel.Quit()

excel_file_path = os.path.abspath("surface.xlsx")
output_folder = "charts"
export_charts_as_png(excel_file_path, output_folder)