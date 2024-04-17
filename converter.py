import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from pathlib import Path

# Improved path handling
xlsx_file_path = Path('important-data.xlsx')
output_folder_path = Path('screenshots')

# Ensure the output directory exists
output_folder_path.mkdir(parents=True, exist_ok=True)

# Read the Excel file into a DataFrame
df = pd.read_excel(xlsx_file_path)

def calculate_column_widths(df, dpi=100, max_total_width=8, padding=2):
    """
    Calculate column widths suitable for matplotlib table,
    setting widths based on the widest content in each column.
    """
    fig, ax = plt.subplots()
    ax.axis('off')
    
    # Create a table cell with the content for each column
    col_widths = []
    total_width = 0  # Initialize total width

    for col in df.columns:
        # Calculate the maximum width for the header
        max_entry_header = len(str(col))
        
        # Calculate the maximum width for the data rows
        max_entry_data = max(df[col].astype(str).apply(len))

        # Choose the maximum width between header and data
        max_entry = max(max_entry_header, max_entry_data)

        cell_text = 'W' * max_entry  # 'W' is typically one of the widest characters
        cell = ax.table(cellText=[[cell_text]], colLabels=[cell_text], loc='center')
        plt.close(fig)
        canvas = FigureCanvas(fig)
        renderer = canvas.get_renderer()
        col_width_pixel = cell.get_window_extent(renderer).width
        col_width_inch = col_width_pixel / dpi
        col_widths.append(col_width_inch + padding / dpi)  # Add padding to each column
        total_width += col_width_inch + padding / dpi

    # Normalize widths to make the total width reasonable
    normalized_widths = [w / total_width for w in col_widths]

    return normalized_widths


def create_table_image(df_slice, image_path, column_widths):
    dpi = 300
    fig_height = len(df_slice) / 5 + 1  # Estimate height based on number of rows, adjust as needed
    fig, ax = plt.subplots(figsize=(8, fig_height), dpi=dpi)  # 8 inches wide figure, adjust as needed
    ax.axis('off')

    # Create table with calculated column widths
    table = ax.table(cellText=df_slice.values, colLabels=df_slice.columns, loc='center', cellLoc='left', colWidths=column_widths)
    table.auto_set_font_size(False)
    table.set_fontsize(10)  # Adjust fontsize as needed
    
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    plt.savefig(image_path, dpi=dpi, bbox_inches='tight', pad_inches=0.05)
    plt.close()

# Set a sample size to speed up the process if the dataframe is very large
# If the dataframe is not too large, you can set this to None to check every row
sample_size = 100

# Calculate optimal column widths from the dataframe sample or the entire dataframe
column_widths = calculate_column_widths(df, sample_size)

# Iterate over the DataFrame in chunks and create images
for start_row in range(0, len(df), 30):
    end_row = min(start_row + 30, len(df))
    df_slice = df.iloc[start_row:end_row]
    image_name = f"df_rows_{start_row}_to_{end_row}.png"
    image_path = output_folder_path / image_name
    create_table_image(df_slice, image_path, column_widths)