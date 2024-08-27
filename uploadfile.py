import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

# Hide the root Tk window
Tk().withdraw()

# Ask the user to select an Excel file
print("Please select the Excel file to sort.")
input_file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

if input_file_path:
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_file_path)

    # Sort the DataFrame by the first column
    df_sorted = df.sort_values(by=df.columns[0])

    # Ask the user where to save the sorted Excel file
    print("Please select the location to save the sorted Excel file.")
    output_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if output_file_path:
        # Save the sorted DataFrame to a new Excel file
        df_sorted.to_excel(output_file_path, index=False)
        print(f"Sorted Excel file has been saved to: {output_file_path}")
    else:
        print("Save operation was cancelled.")
else:
    print("File selection was cancelled.")
