import os
import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
from difflib import SequenceMatcher
 
# Global variable for reference_df
reference_df = None
 
# Dictionary to store date formats for reference columns
reference_date_formats = {}
 
# calculating string similarity ratio
def similarity_ratio(a, b):
    return SequenceMatcher(None, a, b).ratio()
 
# Load reference file
def load_reference_file():
    global reference_df  # Make reference_df global
   
    reference_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    reference_df = pd.read_excel(reference_file_path)
   
    reference_columns_text.config(state=NORMAL)
    reference_columns_text.delete("1.0", END)
    reference_columns_text.insert(END, "Reference File - Columns & Datatypes:\n")
   
    for col, dtype in reference_df.dtypes.items():
        reference_columns_text.insert(END, f"{col}: {dtype}\n")
   
    reference_columns_text.insert(END, f"Total Rows: {len(reference_df)}\n")
    reference_columns_text.config(state=DISABLED)
   
    # Extract date formats for reference columns
    for col in reference_df.columns:
        if reference_df[col].dtype == 'datetime64[ns]':
            unique_dates = reference_df[col].dropna().unique()
            if len(unique_dates) > 0:
                reference_date_formats[col] = unique_dates[0].strftime("%d-%m-%Y")  # Example format, you can adjust it
   
    return reference_file_path
 
# Function to reformat sample file based on reference
def reformat_sample_file(sample_file_path, reference_columns):
    sample_df = pd.read_excel(sample_file_path)
    sample_columns = list(sample_df.columns)
 
    column_mapping = {}
    for reference_col in reference_columns:
        similarity_scores = [similarity_ratio(reference_col, sample_col) for sample_col in sample_columns]
        max_similarity_score = max(similarity_scores)
        if max_similarity_score >= 0.5:
            best_match_index = np.argmax(similarity_scores)
            best_match_col = sample_columns[best_match_index]
            column_mapping[best_match_col] = reference_col
 
    sample_df.rename(columns=column_mapping, inplace=True)
    sample_df = sample_df[reference_columns]
   
    # Preserve non-numeric formats from reference to sample columns
    for col in sample_df.columns:
        reference_dtype = reference_df[col].dtype
        if reference_dtype == 'float64' or reference_dtype == 'int64':
            # Check if reference column contains non-numeric data
            if pd.to_numeric(reference_df[col], errors='coerce').isna().any():
                sample_df[col] = sample_df[col].astype(str)
        elif reference_dtype == 'datetime64[ns]':
            # Check if the reference column has a date format
            if col in reference_date_formats:
                sample_df[col] = pd.to_datetime(sample_df[col], format=reference_date_formats[col], errors='coerce')
   
    return sample_df
# Browse buttons' functions
def browse_reference_file():
    reference_file_path_var.set(load_reference_file())
 
def browse_sample_files_folder():
    sample_files_folder_var.set(filedialog.askdirectory())
 
def browse_output_folder():
    output_folder_var.set(filedialog.askdirectory())
 
# Format files button function
def format_files():
    reference_file_path = reference_file_path_var.get()
    sample_files_folder = sample_files_folder_var.get()
    output_folder = output_folder_var.get()
   
    if not reference_file_path or not sample_files_folder or not output_folder:
        result_label.config(text="Please select all paths.")
        return
   
    try:
        reference_df = pd.read_excel(reference_file_path)
        reference_columns = list(reference_df.columns)
       
        for sample_file_name in os.listdir(sample_files_folder):
            sample_file_path = os.path.join(sample_files_folder, sample_file_name)
            if sample_file_name.endswith(".xlsx"):
                reformatted_df = reformat_sample_file(sample_file_path, reference_columns)
       
                # Save the reformatted file
                output_path = os.path.join(output_folder, sample_file_name)
                reformatted_df.to_excel(output_path, index=False)
       
        result_label.config(text="Formatting completed successfully.")
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")
 
# Create the main window
root = Tk()
root.title("File Formatter")
 
# Reference file path
reference_file_label = Label(root, text="Reference File Path:")
reference_file_label.grid(row=0, column=0, sticky="w")
reference_file_path_var = StringVar()
reference_file_path_button = Button(root, text="Browse", command=browse_reference_file)
reference_file_path_button.grid(row=0, column=1)
reference_file_path_text = Entry(root, textvariable=reference_file_path_var, state=DISABLED, width=50)
reference_file_path_text.grid(row=0, column=2, columnspan=2)
 
# Sample files folder
sample_files_label = Label(root, text="Sample Files Folder:")
sample_files_label.grid(row=1, column=0, sticky="w")
sample_files_folder_var = StringVar()
sample_files_folder_button = Button(root, text="Browse", command=browse_sample_files_folder)
sample_files_folder_button.grid(row=1, column=1)
sample_files_folder_text = Entry(root, textvariable=sample_files_folder_var, state=DISABLED, width=50)
sample_files_folder_text.grid(row=1, column=2, columnspan=2)
 
# Output folder
output_folder_label = Label(root, text="Output Folder:")
output_folder_label.grid(row=2, column=0, sticky="w")
output_folder_var = StringVar()
output_folder_button = Button(root, text="Browse", command=browse_output_folder)
output_folder_button.grid(row=2, column=1)
output_folder_text = Entry(root, textvariable=output_folder_var, state=DISABLED, width=50)
output_folder_text.grid(row=2, column=2, columnspan=2)
 
# Reference file columns display
reference_columns_text = Text(root, height=10, width=50, state=DISABLED)
reference_columns_text.grid(row=3, column=0, columnspan=4)
 
# Format files button
format_button = Button(root, text="Format Files", command=format_files)
format_button.grid(row=4, column=0, columnspan=4)
 
# Result label
result_label = Label(root, text="", fg="green")
result_label.grid(row=5, column=0, columnspan=4)
 
root.mainloop()
 
