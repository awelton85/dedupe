"""
Takes a takeoff worksheet and removes duplicate items, summing the quantities.
"""
import pandas as pd
from tkinter import filedialog

input_file = filedialog.askopenfilename()
output_file = f"{input_file.split('.')[0]}_deduped.xlsx"
df = pd.read_excel(input_file)

# Define custom aggregation functions for each column
agg_functions = {"QTY": "sum"}
for col in df.columns:
    if col not in ["ID", "QTY", "Used"]:
        agg_functions[col] = "first"


# Define custom aggregation function for the 'Where Used' column
def concatenate_unique(series):
    unique_values = series.unique()
    return ", ".join(unique_values)


agg_functions["Used"] = concatenate_unique

# Group by 'ID' and apply the custom aggregation functions, then reset the index
df = df.groupby("ID", as_index=False).agg(agg_functions)

df.to_excel(output_file, index=False)
