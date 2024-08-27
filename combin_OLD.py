import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def select_file(prompt):
    """Prompts the user to select an Excel file and returns the file path."""
    messagebox.showinfo("File Selection", prompt)
    file_path = filedialog.askopenfilename(
        title=prompt,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showwarning("No File Selected",
                               "No file was selected. Exiting.")
        exit()
    return file_path


def save_combined_file(combined_df):
    """Prompts the user to choose a save location for the combined Excel file."""
    messagebox.showinfo("Save Combined File",
                        "Please choose a location to save the combined Excel file.")
    output_file = filedialog.asksaveasfilename(
        title="Save combined file as",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if output_file:
        combined_df.to_excel(output_file, index=False)
        messagebox.showinfo("Success",
                            f"Data has been combined and saved to {output_file}")
    else:
        messagebox.showwarning("Save Cancelled",
                               "The save operation was cancelled.")
        exit()


def main():
    """Main function to handle the file selection, combination, processing, and saving."""
    # Initialize Tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Select the first Excel file
    file1 = select_file("Please select the first Excel file to combine.")

    # Select the second Excel file
    file2 = select_file("Please select the second Excel file to combine.")

    try:
        # Load the data from both Excel files into separate DataFrames
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Combine the data from the two DataFrames
        combined_df = pd.concat([df1, df2], ignore_index=True)

        # Drop the 'Email' column if it exists
        if 'Email' in combined_df.columns:
            combined_df = combined_df.drop(columns=['Email'])

        # Sort the DataFrame by the 'RegNum' column
        if 'RegNum' in combined_df.columns:
            combined_df = combined_df.sort_values(by='RegNum')
        else:
            messagebox.showwarning("Missing Column",
                                   "'RegNum' column is not found. Skipping sort operation.")

        # Save the processed DataFrame to a new Excel file
        save_combined_file(combined_df)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        exit()


if __name__ == "__main__":
    main()
