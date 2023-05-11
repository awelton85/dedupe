"""
Script to remove duplicate items in a takeoff worksheet and sum their quantities.
Updated for python 3.11.3 on 2023-05-11
"""

import pandas as pd
from tkinter import filedialog


def concatenate_unique(series):
    """Function to concatenate unique values from a given pandas series."""
    unique_values = series.unique()
    return ", ".join(unique_values)


# Prompt user to select input file and create output file name
input_file = filedialog.askopenfilename()
output_file = f"{input_file.split('.')[0]}_deduped.xlsx"

# Read input file into a pandas DataFrame
data_frame = pd.read_excel(input_file)

# Define custom aggregation functions for each column
agg_functions = {"QTY": "sum"}
for column in data_frame.columns:
    if column not in ["ID", "QTY", "Used"]:
        agg_functions[column] = "first"

# Add custom aggregation function for the 'Where Used' column
agg_functions["Used"] = concatenate_unique

# Group data by 'ID', apply custom aggregation functions, and reset the index
deduped_data = data_frame.groupby("ID", as_index=False).agg(agg_functions)

# Save the deduplicated data to an output file
deduped_data.to_excel(output_file, index=False)
