import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, simpledialog, Tk
import os

class AddColumnApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Add Column to Multiple Excel Files")
        self.root.geometry("600x800")
        self.root.resizable(True, True)

        # Create instance attributes for file paths and texts
        self.file_paths = [None] * 6
        self.text_entries = [ttk.StringVar() for _ in range(6)]

        # Create GUI components
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        ttk.Label(self.root, text="Add 'Location' Column to Excel Files",
                  font=("Helvetica", 18, "bold"), bootstyle="primary").grid(row=0, column=0, columnspan=3, pady=20)

        # Create Frames for 6 Files
        for i in range(6):
            frame = ttk.Frame(self.root)
            frame.grid(row=i + 1, column=0, columnspan=3, pady=5, padx=10, sticky="ew")

            # File Upload Button
            file_frame = ttk.Frame(frame)
            file_frame.grid(row=0, column=0, padx=(0, 10), sticky="w")
            ttk.Button(file_frame, text=f"Upload File {i + 1}",
                       command=lambda i=i: self.upload_file(i), bootstyle=INFO).pack(side=LEFT)
            label = ttk.Label(file_frame, text="No file selected", bootstyle=SECONDARY, anchor="w")
            label.pack(side=LEFT, fill=X, expand=True)
            setattr(self, f'file{i + 1}_label', label)

            # Combo Box for Column
            column_frame = ttk.Frame(frame)
            column_frame.grid(row=0, column=1, padx=(10, 0), sticky="e")
            ttk.Label(column_frame, text=f"Location {i + 1}:").pack(anchor="w")
            options = ['Maritime lab 232', 'Computer Science lab 206', 'Computer Science lab 404']
            combo_box = ttk.Combobox(column_frame, values=options, textvariable=self.text_entries[i])
            combo_box.pack(fill=X)

        # Add Column Button
        ttk.Button(self.root, text="Add Column to Files",
                   command=self.add_column_to_files, bootstyle=SUCCESS).grid(row=7, column=0, columnspan=3, pady=15, padx=20)

        # Clear and Exit Buttons Frame
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=8, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

        ttk.Button(button_frame, text="Clear", command=self.clear, bootstyle=WARNING).pack(side=LEFT, fill=X, expand=True, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit, bootstyle=DANGER).pack(side=LEFT, fill=X, expand=True, padx=5)

    def upload_file(self, index):
        """Open file dialog to select multiple Excel files and store their paths."""
        files = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if files:
            if len(files) != 6:
                messagebox.showwarning("File Count Error",
                                       "Please select exactly 6 Excel files.")
                return

            for i, file in enumerate(files):
                if i < 6:
                    self.file_paths[i] = file
                    label = getattr(self, f'file{i + 1}_label')
                    label.config(text=os.path.basename(file))
            # Fill any remaining file slots with "No file selected"
            for i in range(len(files), 6):
                self.file_paths[i] = None
                label = getattr(self, f'file{i + 1}_label')
                label.config(text="No file selected")

    def add_column_to_files(self):
        """Add a 'Location' column to each selected file."""
        if None in self.file_paths:
            messagebox.showwarning("Missing Files",
                                   "Please upload all 6 files before adding columns.")
            return

        texts = [text.get() for text in self.text_entries]

        if any(not text for text in texts):
            messagebox.showwarning("Missing Text", "Please select a location for each file.")
            return

        try:
            new_file_name_prefix = simpledialog.askstring("File Name Prefix", "Enter prefix for new files:")
            if not new_file_name_prefix:
                messagebox.showwarning("No Prefix Entered", "You must enter a prefix for the new files.")
                return

            for idx, (file, text) in enumerate(zip(self.file_paths, texts)):
                df = pd.read_excel(file)

                # Add the 'Location' column with the specified text
                df['Location'] = text

                # Save the updated workbook
                output_file = os.path.join(os.path.dirname(file), f"{new_file_name_prefix}_file{idx + 1}_updated.xlsx")
                df.to_excel(output_file, index=False)

            messagebox.showinfo("Success",
                                "Columns added to all files. Updated files have been saved.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def clear(self):
        """Clear file paths, labels, and column entries."""
        for i in range(6):
            self.file_paths[i] = None
            label = getattr(self, f'file{i + 1}_label')
            label.config(text="No file selected")
            self.text_entries[i].set("")

# Example main check to run the app standalone
if __name__ == "__main__":
    root = Tk()
    app = AddColumnApp(root)
    root.mainloop()
