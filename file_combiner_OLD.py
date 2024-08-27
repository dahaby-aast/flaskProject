import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, StringVar, Label, Frame, Tk
import pandas as pd
import os

class FileCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Combiner")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        title_label = ttk.Label(self.root, text="File Combiner",
                               font=("Helvetica", 18, "bold"), bootstyle="primary")
        title_label.grid(row=0, column=0, columnspan=3, pady=20, padx=20, sticky="n")

        # File Selection Frame
        file_frame = ttk.Frame(self.root)
        file_frame.grid(row=1, column=0, columnspan=3, padx=20, pady=10, sticky="ew")

        self.file_paths = [StringVar() for _ in range(6)]
        self.file_labels = []

        for i in range(6):
            frame = ttk.Frame(file_frame)
            frame.grid(row=i, column=0, columnspan=3, pady=5, sticky="ew")

            ttk.Label(frame, text=f"File {i + 1}:").grid(row=0, column=0, padx=10, sticky="w")
            file_label = ttk.Label(frame, text="No file selected", anchor="w")
            file_label.grid(row=0, column=1, sticky="ew")
            self.file_labels.append(file_label)

            ttk.Button(frame, text="Browse", command=lambda i=i: self.browse_file(i), bootstyle=INFO).grid(row=0, column=2, padx=10)

        # Combine Files Button
        combine_button = ttk.Button(self.root, text="Combine Files", command=self.combine_files, bootstyle=SUCCESS)
        combine_button.grid(row=2, column=0, columnspan=3, pady=20, padx=20, sticky="ew")

        # Status Label
        self.status_label = ttk.Label(self.root, text="",
                                      font=("Helvetica", 12, "italic"),
                                      bootstyle="secondary")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

        # Configure column weights to ensure proper expansion
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

    def browse_file(self, index):
        """Open file dialog to select a file and update the corresponding label."""
        files = filedialog.askopenfilenames(
            title=f"Select Excel files",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir="."
        )
        if files:
            if len(files) != 6:
                messagebox.showwarning("File Count Error",
                                       "Please select exactly 6 Excel files.")
                return

            for i, file in enumerate(files):
                if i < 6:
                    self.file_paths[i].set(file)
                    self.file_labels[i].config(text=os.path.basename(file))
            # Fill any remaining file slots with "No file selected"
            for i in range(len(files), 6):
                self.file_paths[i].set("")
                self.file_labels[i].config(text="No file selected")

    def combine_files(self):
        """Combine the selected files into a single Excel file."""
        files = [path.get() for path in self.file_paths]
        files = list(filter(None, files))  # Remove empty strings

        if len(files) != 6:
            messagebox.showwarning("File Count Error",
                                   "Please select exactly 6 files before combining.")
            return

        try:
            combined_df = pd.concat([pd.read_excel(file) for file in files],
                                    ignore_index=True)

            save_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files",
                                                                 "*.xlsx")],
                                                     title="Save Combined File As")
            if save_file:
                combined_df.to_excel(save_file, index=False)
                self.status_label.config(text="Files combined successfully!",
                                         foreground="green")
        except Exception as e:
            messagebox.showerror("Error",
                                 f"An error occurred while combining files: {e}")
            self.status_label.config(text="Failed to combine files.",
                                     foreground="red")

# Create the main window for FileCombinerApp
root = Tk()
app = FileCombinerApp(root)
root.mainloop()
