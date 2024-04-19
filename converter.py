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





import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from pathlib import Path

# Function to calculate column widths
def calculate_column_widths(df, dpi=100, padding=2):
    """
    Calculate column widths suitable for matplotlib table,
    setting widths based on the widest content in each column.
    """
    fig, ax = plt.subplots()
    ax.axis('off')
    
    col_widths = []
    total_width = 0

    for col in df.columns:
        max_entry_header = len(str(col))
        max_entry_data = max(df[col].astype(str).apply(len))
        max_entry = max(max_entry_header, max_entry_data)

        cell_text = 'W' * max_entry
        cell = ax.table(cellText=[[cell_text]], colLabels=[cell_text], loc='center')
        plt.close(fig)
        canvas = FigureCanvas(fig)
        renderer = canvas.get_renderer()
        col_width_pixel = cell.get_window_extent(renderer).width
        col_width_inch = col_width_pixel / dpi
        col_widths.append(col_width_inch + padding / dpi)
        total_width += col_width_inch + padding / dpi

    normalized_widths = [w / total_width for w in col_widths]
    return normalized_widths

# Function to create and save table images
def create_table_image(df_slice, image_path, column_widths):
    dpi = 300
    fig_height = len(df_slice) / 5 + 1
    fig, ax = plt.subplots(figsize=(8, fig_height), dpi=dpi)
    ax.axis('off')

    table = ax.table(cellText=df_slice.values, colLabels=df_slice.columns, loc='center', cellLoc='left', colWidths=column_widths)
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    plt.savefig(image_path, dpi=dpi, bbox_inches='tight', pad_inches=0.05)
    plt.close()

# Directories setup
xlsx_folder_path = Path('xlsx_files')
output_folder_path = Path('screenshots')
output_folder_path.mkdir(parents=True, exist_ok=True)

# Iterate over each Excel file in the "xlsx_files" directory
for xlsx_file in xlsx_folder_path.glob('*.xlsx'):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(xlsx_file)
    
    # Calculate optimal column widths from the dataframe
    column_widths = calculate_column_widths(df)

    # Since each file might contain different ranges, process the entire DataFrame
    image_name = f"{xlsx_file.stem}.png"
    image_path = output_folder_path / image_name
    create_table_image(df, image_path, column_widths)
    