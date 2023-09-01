import os
import pandas as pd
import numpy as np
from difflib import SequenceMatcher

# calculating string similarity ratio
def similarity_ratio(a, b):
    return SequenceMatcher(None, a, b).ratio()

# Load reference file
reference_file_path = "path/to/reference file"
reference_df = pd.read_excel(reference_file_path)
#print(reference_df)
reference_columns = list(reference_df.columns)
 
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
    # Apply exact data types from reference to sample columns
    for col in sample_df.columns:
        reference_dtype = reference_df[col].dtype
        if reference_dtype == 'float64' or reference_dtype == 'int64':
            sample_df[col] = pd.to_numeric(sample_df[col], errors='coerce')
            sample_df[col] = sample_df[col].fillna(0)
        elif reference_dtype == 'datetime64[ns]':
            sample_df[col] = pd.to_datetime(sample_df[col], errors='coerce')
        else:
            sample_df[col] = sample_df[col].astype(reference_dtype)
 
    return sample_df
 
# Folder containing sample files (folder containing files that need to be formatted)
sample_files_folder = "path/to/folder containing sample files"

# Folder containing output files
output_folder = "path/to/output folder"
 
# Process each sample file in the folder
for sample_file_name in os.listdir(sample_files_folder):
    sample_file_path = os.path.join(sample_files_folder, sample_file_name)
    if sample_file_name.endswith(".xlsx"):
        reformatted_df = reformat_sample_file(sample_file_path, reference_columns)
 
        # Save the reformatted file
        output_path = os.path.join(output_folder, sample_file_name)
        reformatted_df.to_excel(output_path, index=False)
