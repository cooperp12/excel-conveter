Generating Table Images from Excel Data

This script generates images of data tables from an Excel file using matplotlib in Python. It's particularly useful when you want to visualize data tables and share them in image format.

Requirements

- Python 3.x
- pandas
- matplotlib

You can install the required dependencies using pip:

pip install pandas matplotlib

Usage

1. Place your Excel file (important-data.xlsx) in the same directory as this script.
2. Ensure you have a folder named screenshots where the generated images will be saved.
3. Run the script.

The script will read the Excel file, calculate the optimal column widths for the table, and generate images of the data tables in chunks of rows. Images will be saved in the screenshots folder.

Parameters

- xlsx_file_path: Path to the Excel file.
- output_folder_path: Path to the folder where images will be saved.
- sample_size: Set a sample size to speed up the process if the dataframe is very large. If the dataframe is not too large, you can set this to None to check every row.
- dpi: Dots per inch for image resolution.
- max_total_width: Maximum total width of the table.
- padding: Padding between columns in inches.

Customization

You can adjust the following parameters in the script according to your needs:

- dpi: Adjust image resolution.
- fig_height: Estimate the figure height based on the number of rows in the data table.
- figsize: Adjust the figure size for the generated images.
- table.set_fontsize(): Set the font size for the table.

Example

For a sample size of 100 rows, the script will generate images of data tables in chunks of 30 rows each, named df_rows_<start_row>_to_<end_row>.png, where <start_row> and <end_row> represent the row indices.

python

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