import pandas as pd
from pathlib import Path
import os
import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import argrelextrema

def plot_sequences(x1, y1, x2, y2, label1, label2, save_path):
    # Plot two sequences for comparison and save the plot
    plt.figure(figsize=(10, 6))
    plt.plot(x1, y1, label=label1)
    plt.plot(x2, y2, label=label2)
    plt.xlabel('Wavelength (nm)')
    plt.ylabel('Reflectance')
    plt.legend()
    plt.title(f'Comparison of Sequences: {label1} vs {label2}')
    plt.savefig(save_path)
    plt.close()

def calculate_slope(x, y):
    # Calculate the slope between consecutive points
    return pd.Series(y).diff() / pd.Series(x).diff()

def find_local_minima(x, y):
    # Find local minima in the sequence
    indices = argrelextrema(np.array(y), np.less)[0]
    return x[indices], y[indices]

def filter_files_by_minima(folder, x_array, tolerance, required_matches):
    # Filter files based on the occurrence of local minima within a tolerance range
    files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
    qualifying_files = []
    for file in files:
        df = pd.read_excel(os.path.join(folder, file))
        minima_x, _ = find_local_minima(df.iloc[:, 0].values, df.iloc[:, 1].values)
        if len(minima_x) < 60 and sum(any(abs(minima - x) <= tolerance for minima in minima_x) for x in x_array) >= required_matches:
            qualifying_files.append(file)
    return qualifying_files

def save_top_files_to_excel(difference_dfs, save_path):
    # Save the top files to an Excel file
    data = {
        'File Name': [fn[:-5] for _, fn in difference_dfs[:20]],
        'Count': [count for count, _ in difference_dfs[:20]],
        'Match %': [round((count / 2101) * 100, 2) for count, _ in difference_dfs[:20]]
    }
    pd.DataFrame(data).to_excel(save_path, index=False)

# Path to the reference Excel file
file1 = Path('reference_file.xlsx')
# Folder containing the files to compare
folder = Path('data_folder')
# Folder to save the plots
plots_folder = Path('plots')
# Array of x values to compare local minima
x_array = [380, 1160, 1407, 1918, 2196, 2260, 2352]
# Tolerance for local minima matching
tolerance = 10
# Minimum number of matches required
required_matches = 4

# Read the reference file
main_df = pd.read_excel(file1)
main_df['slope'] = calculate_slope(main_df.iloc[:, 0], main_df.iloc[:, 1])
difference_dfs = []

# Get files that meet the criteria based on local minima
qualifying_files = filter_files_by_minima(folder, x_array, tolerance, required_matches)

# Process each qualifying file
for file in qualifying_files:
    df = pd.read_excel(os.path.join(folder, file))
    df['slope'] = calculate_slope(df.iloc[:, 0], df.iloc[:, 1])
    df['difference'] = abs(df['slope'] - main_df['slope'])
    count = ((df['difference'] >= 0) & (df['difference'] <= 0.001)).sum()
    difference_dfs.append((count, file))
difference_dfs.sort(reverse=True, key=lambda x: x[0])

# Output top files and generate plots
if difference_dfs:
    print("Top 10 files with the highest count values:")
    for count, file in difference_dfs[:10]:
        print(f"File: {file[:-5]}, Count: {count}")
        df2 = pd.read_excel(os.path.join(folder, file))
        plot_sequences(main_df.iloc[:, 0], main_df.iloc[:, 1], df2.iloc[:, 0], df2.iloc[:, 1], file1.stem, file[:-5],
                       plots_folder.joinpath(f'{file1.stem}_vs_{file[:-5]}.png'))

    save_top_files_to_excel(difference_dfs, plots_folder.joinpath(f'{file1.stem}_Top_20_Files.xlsx'))
else:
    print("No files met the criteria.")
