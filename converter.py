import xlwings as xw
import os

# Create a folder for the new Excel files if it doesn't exist
xlsx_folder = "xlsx_files"
if not os.path.exists(xlsx_folder):
    os.makedirs(xlsx_folder)

# Specify the path to your original Excel file
xlsx_file_path = 'important-data.xlsx'

# Define the number of rows per saved range
rows_per_range = 30

# Start an instance of Excel and open the workbook
app = xw.App(visible=True)
wb = app.books.open(xlsx_file_path)
sheet = wb.sheets[0]  # Assumes data is in the first sheet

# Calculate how many ranges (and thus files) we need to create
row_count = sheet.range('A1').end('down').row
range_count = (row_count + rows_per_range - 1) // rows_per_range

# Loop over each range
for i in range(range_count):
    # Calculate the range for this segment
    start_row = i * rows_per_range + 1
    end_row = min((i + 1) * rows_per_range, row_count)
    
    # Create a new workbook for this range
    new_wb = app.books.add()
    new_sheet = new_wb.sheets[0]
    
    # Copy the range from the original sheet and paste it into the new sheet
    sheet.range(f'A{start_row}:D{end_row}').api.Copy(Destination=new_sheet.range('A1').api)
    
    # Save the new workbook with an informative filename
    new_wb.save(os.path.join(xlsx_folder, f'range_{start_row}_to_{end_row}.xlsx'))
    new_wb.close()

# Quit Excel
wb.close()
app.quit()




