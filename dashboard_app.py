import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import Tk, Toplevel, filedialog, messagebox, simpledialog, Listbox
import os
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=SyntaxWarning)
class DashboardApp:
    def __init__(self, root):
        self.root = root
        self.current_theme = "cosmo"  # Default theme
        self.root.title("Dashboard")
        self.root.geometry("400x450")
        self.root.resizable(True, True)
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        ttk.Label(self.root, text="Main Dashboard",
                  font=("Helvetica", 18, "bold"), bootstyle="primary").pack(pady=20)

        # Button to open AddColumnApp
        ttk.Button(self.root, text="Open Add Column App",
                   command=self.open_add_column_app, bootstyle=SUCCESS).pack(pady=10, padx=20, fill="x")

        # Button to open FileCombinerApp
        ttk.Button(self.root, text="Open File Combiner App",
                   command=self.open_file_combiner_app, bootstyle=INFO).pack(pady=10, padx=20, fill="x")

        # Button to open a file
        ttk.Button(self.root, text="Open File",
                   command=self.open_file, bootstyle=PRIMARY).pack(pady=10, padx=20, fill="x")

        # Theme Toggle Button
        ttk.Button(self.root, text="Dark Mode", command=self.toggle_theme,
                   bootstyle=WARNING).pack(pady=10, padx=20, fill="x")

        # Change Theme Color Button
        ttk.Button(self.root, text="Change Theme Color",
                   command=self.change_theme_color, bootstyle=INFO).pack(pady=10, padx=20, fill="x")

        # Exit Button
        ttk.Button(self.root, text="Exit", command=self.root.quit,
                   bootstyle=DANGER).pack(pady=10, padx=20, fill="x")

    def open_add_column_app(self):
        self.root.iconify()  # Minimize the dashboard
        top = Toplevel(self.root)
        top.title("Add Column App")
        top.geometry("800x600")
        AddColumnApp(top)
        top.protocol("WM_DELETE_WINDOW", lambda: self.on_close_app(top))

    def open_file_combiner_app(self):
        self.root.iconify()  # Minimize the dashboard
        top = Toplevel(self.root)
        top.title("File Combiner App")
        top.geometry("800x600")
        FileCombinerApp(top)
        top.protocol("WM_DELETE_WINDOW", lambda: self.on_close_app(top))

    def open_file(self):
        file_path = filedialog.askopenfilename(
            title="Select a File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            try:
                # Open the Excel file with the default application
                os.startfile(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def on_close_app(self, window):
        window.destroy()
        self.root.deiconify()  # Restore t
        # he dashboard

    def toggle_theme(self):
        # Toggle between light and dark mode
        new_theme = "darkly" if self.current_theme == "cosmo" else "cosmo"
        self.root.style.theme_use(new_theme)
        self.current_theme = new_theme

    def change_theme_color(self):
        # Example to change to a specific theme
        new_theme = "minty"  # or any other available theme
        self.root.style.theme_use(new_theme)


class AddColumnApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Add Column to Multiple Excel Files")
        self.root.geometry("600x800")
        self.root.resizable(True, True)
        self.file_paths = [None] * 6
        self.text_entries = [ttk.StringVar() for _ in range(6)]
        self.records_labels = [None] * 6  # To hold labels for record counts
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        ttk.Label(self.root, text="Add 'Location' Column to Excel Files",
                  font=("Helvetica", 18, "bold"), bootstyle="primary").grid(row=0, column=0, columnspan=3, pady=20)

        for i in range(6):
            frame = ttk.Frame(self.root, borderwidth=2, relief="flat")
            frame.grid(row=i + 1, column=0, columnspan=3, pady=5, padx=10, sticky="ew")

            file_frame = ttk.Frame(frame)
            file_frame.grid(row=0, column=0, padx=(0, 10), sticky="w")
            ttk.Button(file_frame, text=f"Upload File {i + 1}",
                       command=lambda i=i: self.upload_file(i), bootstyle=INFO).pack(side=LEFT)
            label = ttk.Label(file_frame, text="No file selected", bootstyle=SECONDARY, anchor="w")
            label.pack(side=LEFT, fill="x", expand=True)
            setattr(self, f'file{i + 1}_label', label)

            # Create labels for record counts
            self.records_labels[i] = ttk.Label(file_frame, text="Records: N/A", bootstyle=SECONDARY, anchor="w")
            self.records_labels[i].pack(side=LEFT, padx=(10, 0))

            column_frame = ttk.Frame(frame)
            column_frame.grid(row=0, column=1, padx=(10, 0), sticky="e")
            ttk.Label(column_frame, text=f"Location {i + 1}:").pack(anchor="w")
            options = ['Maritime_lab 232', 'CS_lab 206', 'CS_lab 404']
            combo_box = ttk.Combobox(column_frame, values=options, textvariable=self.text_entries[i])
            combo_box.pack(fill="x")

        ttk.Button(self.root, text="Add Column to Files",
                   command=self.add_column_to_files, bootstyle=SUCCESS).grid(row=7, column=0, columnspan=3, pady=15, padx=20)

        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=8, column=0, columnspan=3, pady=10, padx=20, sticky="ew")
        ttk.Button(button_frame, text="Clear", command=self.clear, bootstyle=WARNING).pack(side=LEFT, fill="x", expand=True, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.destroy, bootstyle=DANGER).pack(side=LEFT, fill="x", expand=True, padx=5)

    def upload_file(self, index):
        files = filedialog.askopenfilenames(title="Select Excel Files",
                                            filetypes=[("Excel files", "*.xlsx *.xls")])
        if files:
            if len(files) != 6:
                messagebox.showwarning("File Count Error", "Please select exactly 6 Excel files.")
                return

            for i, file in enumerate(files):
                if i < 6:
                    self.file_paths[i] = file
                    label = getattr(self, f'file{i + 1}_label')
                    label.config(text=os.path.basename(file))
                    self.update_record_count(i)  # Update record count for each file
            for i in range(len(files), 6):
                self.file_paths[i] = None
                label = getattr(self, f'file{i + 1}_label')
                label.config(text="No file selected")
                self.records_labels[i].config(text="Records: N/A")  # Reset record count label

    def update_record_count(self, index):
        try:
            df = pd.read_excel(self.file_paths[index])
            record_count = len(df)
            self.records_labels[index].config(text=f"Records: {record_count}")
        except Exception as e:
            self.records_labels[index].config(text="Records: Error")
            print(f"Error reading {self.file_paths[index]}: {e}")

    def add_column_to_files(self):
        if None in self.file_paths:
            messagebox.showwarning("Missing Files", "Please upload all 6 files before adding columns.")
            return

        texts = [text.get() for text in self.text_entries]

        if any(not text for text in texts):
            messagebox.showwarning("Missing Text", "Please select a location for each file.")
            return

        try:
            new_file_name_prefix = simpledialog.askstring("File Name Prefix",
                                                          "Enter prefix for new files:")
            if not new_file_name_prefix:
                messagebox.showwarning("No Prefix Entered", "You must enter a prefix for the new files.")
                return

            for idx, (file, text) in enumerate(zip(self.file_paths, texts)):
                df = pd.read_excel(file)
                df['Location'] = text
                output_file = os.path.join(os.path.dirname(file),
                                           f"{new_file_name_prefix}_file{idx + 1}_updated.xlsx")
                df.to_excel(output_file, index=False)

            messagebox.showinfo("Success", "Columns added to all files. Updated files have been saved.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def clear(self):
        self.file_paths = [None] * 6
        for i in range(6):
            label = getattr(self, f'file{i + 1}_label')
            label.config(text="No file selected")
            self.records_labels[i].config(text="Records: N/A")
            self.text_entries[i].set("")


class FileCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Combiner")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        self.file_paths = []
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        ttk.Label(self.root, text="Combine Excel Files",
                  font=("Helvetica", 18, "bold"), bootstyle="primary").grid(row=0, column=0, columnspan=3, pady=20)

        ttk.Button(self.root, text="Upload Files", command=self.upload_files, bootstyle=INFO).grid(
            row=1, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

        self.file_listbox = Listbox(self.root, height=10)
        self.file_listbox.grid(row=2, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

        ttk.Button(self.root, text="Combine Files", command=self.combine_files, bootstyle=SUCCESS).grid(
            row=3, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

        ttk.Button(self.root, text="Clear", command=self.clear, bootstyle=WARNING).grid(
            row=4, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

        ttk.Button(self.root, text="Exit", command=self.root.destroy, bootstyle=DANGER).grid(
            row=5, column=0, columnspan=3, pady=10, padx=20, sticky="ew")

    def upload_files(self):
        files = filedialog.askopenfilenames(title="Select Excel Files",
                                            filetypes=[("Excel files", "*.xlsx *.xls")])
        if files:
            self.file_paths = list(files)
            self.file_listbox.delete(0, 'end')
            for file in files:
                self.file_listbox.insert('end', os.path.basename(file))

    def combine_files(self):
        if not self.file_paths:
            messagebox.showwarning("No Files", "Please upload some Excel files before combining.")
            return

        try:
            combined_df = pd.concat([pd.read_excel(file) for file in self.file_paths])
            combined_df = combined_df.sort_values(by='RegNum')
            combined_df = combined_df.drop(columns=['Email'], errors='ignore')

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx *.xls")],
                                                     title="Save Combined File")
            if save_path:
                combined_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", "Files combined and saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def clear(self):
        self.file_paths = []
        self.file_listbox.delete(0, 'end')

if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = DashboardApp(root)
    root.mainloop()