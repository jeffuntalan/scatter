# BioAssets Scatter Plot Viewer and Maker for Pre-formatted Excel Files
# Description: This script reads data from Excel files and displays scatter plots for each sheet.

import matplotlib.pyplot as plt
import pandas as pd
import glob
import os

# Function to read the title and data series from a sheet
def read_data_series_from_sheet(sheet):
    # Get title from the first cell
    title = sheet.columns[0]
    sheet.columns = sheet.iloc[0]
    sheet = sheet[1:]
    
    # Reset the DataFrame index
    sheet.reset_index(drop=True, inplace=True)
    
    # Convert the DataFrame columns to the appropriate data format
    data = {
        'T1': sheet['T1'].dropna().astype(float).tolist(),
        'T2': sheet['T2'].dropna().astype(float).tolist(),
        'T3': sheet['T3'].dropna().astype(float).tolist()
    }
    return title, data

# Function to process a single sheet and generate a plot
def process_sheet(sheet, title, output_path):
    # Read the title and data series from the sheet
    title, data = read_data_series_from_sheet(sheet)

    # Calculate average for each treatment
    averages = {key: sum(values) / len(values) for key, values in data.items()}

    # Create scatter plot
    x_labels = ['T1', 'T2', 'T3']
    for i, treatment in enumerate(x_labels):
        values = data[treatment]
        x = [i + 1] * len(values)  # X-axis positions for each treatment group
        color = 'black' if treatment == 'T1' else 'blue' if treatment == 'T2' else 'red'  # Color based on treatment
        plt.scatter(x, values, label=f'{treatment}', alpha=0.5, color=color)

        # Plot average line for each treatment
        plt.hlines(averages[treatment], i + 0.8, i + 1.2, colors=color, linestyles='dashed', label=f'{treatment} avg')

    # Customize plot
    plt.xticks([1, 2, 3], x_labels)  # Set x-axis labels
    plt.xlabel('Treatment')
    plt.ylabel('Fold Change')
    plt.title(title)
    plt.legend(loc=2)
    plt.tight_layout()

    # Save the plot to the specified path
    plt.savefig(output_path)
    plt.close()

# Function to process a single file and generate plots for each sheet
def process_file(file_path, output_dir):
    # Check the file extension
    if file_path.endswith('.xlsx'):
        sheets = pd.read_excel(file_path, sheet_name=['Day 14', 'Day 28'])
    else:
        raise ValueError("Unsupported file format")
    
    # Process each sheet
    base_filename = os.path.splitext(os.path.basename(file_path))[0]
    for sheet_name, sheet in sheets.items():
        output_path = os.path.join(output_dir, f"{base_filename}_{sheet_name.replace(' ', '_')}.png")
        process_sheet(sheet, sheet_name, output_path)

# List of Excel files to process
files = glob.glob('scatter*.xlsx')

# Create output directory if it does not exist
output_dir = 'scatters'
os.makedirs(output_dir, exist_ok=True)

# Process each file
for file_path in files:
    process_file(file_path, output_dir)
