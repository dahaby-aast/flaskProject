import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename

# Hide the root Tk window
Tk().withdraw()

try:
    # Ask the user to select an Excel file
    input_file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if input_file_path:
        # Load the Excel file into a DataFrame

        df = pd.read_excel(input_file_path)
        if 'Email' in df.columns:
            # Delete the 'Email' column
            df.drop(columns=['Email'], inplace=True)
        else:
            messagebox.showwarning("Warning", "'email' column not found.")
        # Sort the DataFrame by the first column
        df_sorted = df.sort_values(by=df.columns[0])


        # Ask the user where to save the sorted Excel file
        output_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if output_file_path:
            # Save the sorted DataFrame to a new Excel file
            df_sorted.to_excel(output_file_path, index=False)
            messagebox.showinfo("Follla keda", f"Sorted Excel  saved tamam gedan :\n{output_file_path}")
        else:
            messagebox.showwarning("fe 7aga 3'alat", "7ot esm lel File ya 3amona.")
    else:
        messagebox.showwarning("bala7", "Mate7'tar el file ya 3amooona.")
except Exception as e:
    # Display any exceptions that occur in a message box
    messagebox.showerror("oppppppa", f"fe moooosiba ya 3amona:\n{str(e)}")
