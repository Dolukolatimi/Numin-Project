import os
import pandas as pd

print("Merging Excel files...")

def read_excel_files(directory='Train_excel', out_xlsx='merged_data.xlsx'):

    def _is_unnamed(colname):
        """Check if a column name is auto-generated like 'Unnamed: 0' or 'Untitled'."""
        return isinstance(colname, str) and (colname.lower().startswith("unnamed") or colname.lower().startswith("untitled"))

    all_frames = []
    no_files_with_missing = 0
    # Get all Excel files in the directory
    files = [f for f in os.listdir(directory) if f.lower().endswith(('.xlsx', '.xls'))]

    if not files:
        print(f"No Excel files found in '{directory}'.")
        return pd.DataFrame()

    # Loop through all Excel files
    for filename in files:
        path = os.path.join(directory, filename)
        try:
            xls = pd.ExcelFile(path)  # Read workbook
        except Exception as e:
            print(f"Skip file {filename}: {e}")
            continue

        # Loop through all sheets in the file
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet)
            except Exception as e:
                print(f"Skip sheet {sheet} in {filename}: {e}")
                continue

            # Skip if sheet has no columns
            if df is None or df.shape[1] == 0:
                continue

            # Drop rows that are entirely empty
            df = df.dropna(how='all')
            if df.empty:
                continue

            # Check for missing/NaN values in the dataframe
            if df.isnull().values.any():
                no_files_with_missing += 1

            # Rename the first column if it's an auto-generated header
            if _is_unnamed(df.columns[0]):
                df.rename(columns={df.columns[0]: 'days'}, inplace=True)

            # Add file and sheet information for traceability
            # df.insert(0, 'source_file', filename)
            # df.insert(1, 'source_sheet', sheet)
            all_frames.append(df)

    # Stop if nothing was collected
    if not all_frames:
        print("No data found to merge.")
        return pd.DataFrame()
    
    # Check if any files had missing values
    if no_files_with_missing > 0:
        print(f"Warning: {no_files_with_missing} sheets contained missing/NaN values.")

    # Combine all sheets into one DataFrame
    merged = pd.concat(all_frames, ignore_index=True)

    # Save outputs in Excel format
    merged.to_excel(out_xlsx, index=False)
    print(f"Merged {len(all_frames)} sheets from {len(files)} files -> {out_xlsx}")
    
    return merged

if __name__ == "__main__":
    read_excel_files()
    print("Merged Excel files.")
