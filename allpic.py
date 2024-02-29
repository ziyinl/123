import os
import pydicom
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import math
import numpy as np
import pandas as pd

def read_dicom_and_plot(filepath):
    # Read the DICOM file
    ds = pydicom.dcmread(filepath)

    # Determine number of columns and rows for plotting
    numcols = 10
    numrows = math.ceil(ds.NumberOfFrames / numcols)

    # Create a plot
    figure = plt.figure(figsize=(20, 16))
    plt.title('All fig')
    cmap = mpl.cm.jet
    grid = plt.GridSpec(numrows, numcols, wspace=0.1, hspace=0.1)

    # Plot each frame in DICOM file
    for i in range(ds.NumberOfFrames):
        tempimage = ds.pixel_array[i, :, :]
        plt.subplot(grid[i // numcols, i % numcols])
        dsarray = tempimage.reshape(-1)
        dsthreshold = np.percentile(dsarray, 99.9)
        norm = mpl.colors.Normalize(vmin=0, vmax=dsthreshold)
        plt.imshow(tempimage, cmap=cmap, norm=norm)

    plt.show(block=False)  # Show the plot without blocking the code

def collect_user_input(filename):
    # Prompt user input for data entry
    start_study = input(f"請輸入檔案 {filename} 的起始研究編號：")
    end_study = input(f"請輸入檔案 {filename} 的結束研究編號：")

    return {'檔名': filename, '開始': start_study, '結束': end_study}

# Path to the directory containing DICOM files
dicom_directory = r'/Users/ziyinlin/Library/CloudStorage/OneDrive-東海大學/論文/trodatall/PD'

# Create an empty list to store data
data_list = []

# Read DICOM files and plot images
for filename in os.listdir(dicom_directory):
    if filename.endswith(''):
        filepath = os.path.join(dicom_directory, filename)
        read_dicom_and_plot(filepath)
        user_input = collect_user_input(filename)
        data_list.append(user_input)

# Convert list of dictionaries to DataFrame
df = pd.DataFrame(data_list)

# Save DataFrame to Excel
excel_filename = 'dicom_datatrue.xlsx'
df.to_excel(excel_filename, index=False)
print(f"Excel file '{excel_filename}' saved successfully.")
