import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import Tk, Toplevel, filedialog, messagebox, simpledialog, Listbox
import sys
import os
import pandas as pd
import warnings
from datetime import datetime

warnings.filterwarnings("ignore", category=SyntaxWarning)

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


class DashboardApp:
    def __init__(self, root):
        self.root = root
        self.style = ttk.Style()  # Create a ttkbootstrap Style object
        self.current_theme = "darkly"
        self.style.theme_use(self.current_theme)  # Apply the initial theme
        self.root.title("Dashboard")
        self.root.geometry("400x720")
        self.root.resizable(True, True)
        self.create_widgets()

    def create_widgets(self):
        self.logo = ttk.PhotoImage(file=resource_path("uploads/logo1.png"))
        ttk.Label(self.root, image=self.logo).pack(pady=10)
        ttk.Label(self.root, text="My Dashboard", font=("Helvetica", 18, "bold"), bootstyle="primary").pack(pady=20)

        current_date = datetime.now().strftime("%Y-%m-%d")
        date_frame = ttk.Frame(self.root, bootstyle="info")
        date_frame.pack(pady=10, padx=20, fill="x")

        self.date_label = ttk.Label(date_frame, text=f"Date: {current_date}", font=("Helvetica", 14, "bold"),
                                    bootstyle="light", background="#303030", foreground="#fdfdfd")
        self.date_label.pack(padx=10, pady=5)

        time_frame = ttk.Frame(self.root, bootstyle="info")
        time_frame.pack(pady=10, padx=20, fill="x")
        self.time_label = ttk.Label(time_frame, text="", font=("Helvetica", 14, "bold"), bootstyle="light", foreground="#fdfdfd")
        self.time_label.pack(padx=10, pady=5)

        self.update_time()

        ttk.Button(self.root, text="Open Add Column App", command=self.open_add_column_app, bootstyle=SUCCESS).pack(pady=10, padx=20, fill="x")
        ttk.Button(self.root, text="Open File Combiner App", command=self.open_file_combiner_app, bootstyle=INFO).pack(pady=10, padx=20, fill="x")
        ttk.Button(self.root, text="Open File", command=self.open_file, bootstyle=PRIMARY).pack(pady=10, padx=20, fill="x")
        ttk.Button(self.root, text="Dark Mode", command=self.toggle_theme, bootstyle=WARNING).pack(pady=10, padx=20, fill="x")
        ttk.Button(self.root, text="Change Theme Color", command=self.change_theme_color, bootstyle=INFO).pack(pady=10, padx=20, fill="x")
        ttk.Button(self.root, text="Exit", command=self.root.quit, bootstyle=DANGER).pack(pady=10, padx=20, fill="x")

        ttk.Label(self.root, text="Все права защищены @ Дахаби 2024", font=("Helvetica", 13, "bold"),
                  bootstyle="secondary", anchor="w", padding=(10, 5), foreground="#292f35").pack(side="bottom", fill="x")

    def update_time(self):
        self.time_label.config(text=f"Time: {datetime.now().strftime('%H:%M:%S')}")
        self.root.after(1000, self.update_time)

    def open_add_column_app(self):
        self._open_app_window("Add Column App", AddColumnApp)

    def open_file_combiner_app(self):
        self._open_app_window("File Combiner App", FileCombinerApp)

    def open_file(self):
        file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if file_path:
            try:
                os.startfile(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def _open_app_window(self, title, app_class):
        self.root.iconify()
        top = Toplevel(self.root)
        top.title(title)
        top.geometry("800x600")
        app_class(top)
        top.protocol("WM_DELETE_WINDOW", lambda: self.on_close_app(top))

    def on_close_app(self, window):
        window.destroy()
        self.root.deiconify()

    def toggle_theme(self):
        self.current_theme = "darkly" if self.current_theme == "cosmo" else "cosmo"
        self.style.theme_use(self.current_theme)  # Use the style object to change theme

    def change_theme_color(self):
        self.style.theme_use("minty")  # Change theme using the style object


class AddColumnApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Add Column to Multiple Excel Files")
        self.root.geometry("700x800")
        self.root.resizable(True, True)
        self.file_paths = [None] * 6
        self.text_entries = [ttk.StringVar() for _ in range(6)]
        self.records_labels = [None] * 6
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self.root, text="Add 'Location' Column to Excel Files", font=("Helvetica", 18, "bold"), bootstyle="primary").grid(row=0, column=0, columnspan=3, pady=20)

        for i in range(6):
            frame = ttk.Frame(self.root, borderwidth=2, relief="flat")
            frame.grid(row=i + 1, column=0, columnspan=3, pady=5, padx=10, sticky="ew")

            file_frame = ttk.Frame(frame)
            file_frame.grid(row=0, column=0, padx=(0, 10), sticky="w")
            ttk.Button(file_frame, text=f"Upload File {i + 1}", command=lambda i=i: self.upload_file(i), bootstyle=INFO).pack(side=LEFT)
            label = ttk.Label(file_frame, text="No file selected", bootstyle=SECONDARY, anchor="w")
            label.pack(side=LEFT, fill="x", expand=True)
            setattr(self, f'file{i + 1}_label', label)

            self.records_labels[i] = ttk.Label(file_frame, text="Records: N/A", bootstyle=SECONDARY, anchor="w")
            self.records_labels[i].pack(side=LEFT, padx=(10, 0))

            column_frame = ttk.Frame(frame)
            column_frame.grid(row=0, column=1, padx=(10, 0), sticky="e")
            ttk.Label(column_frame, text=f"Location {i + 1}:").pack(anchor="w")
            options = ['Maritime_lab 232', 'CS_lab 206', 'CS_lab 404']
            combo_box = ttk.Combobox(column_frame, values=options, textvariable=self.text_entries[i])
            combo_box.pack(fill="x")

        ttk.Button(self.root, text="Add Column to Files", command=self.add_column_to_files, bootstyle=SUCCESS).grid(row=7, column=0, columnspan=3, pady=15, padx=20)
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=8, column=0, columnspan=3, pady=10, padx=20, sticky="ew")
        ttk.Button(button_frame, text="Clear", command=self.clear, bootstyle=WARNING).pack(side=LEFT, fill="x", expand=True, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.destroy, bootstyle=DANGER).pack(side=LEFT, fill="x", expand=True, padx=5)

    def upload_file(self, index):
        files = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel files", "*.xlsx *.xls")])
        if files:
            for i in range(min(len(files), 6)):
                self.file_paths[i] = files[i]
                getattr(self, f'file{i + 1}_label').config(text=os.path.basename(files[i]))
                self.update_record_count(i)
            for i in range(len(files), 6):
                self.file_paths[i] = None
                getattr(self, f'file{i + 1}_label').config(text="No file selected")
                self.records_labels[i].config(text="Records: N/A")

    def update_record_count(self, index):
        try:
            df = pd.read_excel(self.file_paths[index])
            self.records_labels[index].config(text=f"Records: {len(df)}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reading {os.path.basename(self.file_paths[index])}: {e}")

    def add_column_to_files(self):
        if None in self.file_paths:
            messagebox.showwarning("Incomplete Selection", "Please upload all files before proceeding.")
            return

        for i, file_path in enumerate(self.file_paths):
            if file_path:
                try:
                    df = pd.read_excel(file_path)
                    df['Location'] = self.text_entries[i].get()
                    output_file = f"{os.path.splitext(file_path)[0]}_updated.xlsx"
                    df.to_excel(output_file, index=False)
                    messagebox.showinfo("Success", f"Column added successfully to {os.path.basename(file_path)}.")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred while processing {os.path.basename(file_path)}: {e}")

    def clear(self):
        self.file_paths = [None] * 6
        for i in range(6):
            getattr(self, f'file{i + 1}_label').config(text="No file selected")
            self.text_entries[i].set("")
            self.records_labels[i].config(text="Records: N/A")


class FileCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Combiner")
        self.root.geometry("700x800")
        self.root.resizable(True, True)
        self.file_paths = []
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self.root, text="Combine Multiple Excel Files", font=("Helvetica", 18, "bold"), bootstyle="primary").grid(row=0, column=0, columnspan=3, pady=20)

        upload_button = ttk.Button(self.root, text="Upload Files", command=self.upload_files, bootstyle=INFO)
        upload_button.grid(row=1, column=0, columnspan=3, pady=10, padx=20)

        self.files_listbox = Listbox(self.root, selectmode="multiple", font=("Helvetica", 10), width=70, height=10)
        self.files_listbox.grid(row=2, column=0, columnspan=3, pady=10, padx=20)

        ttk.Button(self.root, text="Combine Files", command=self.combine_files, bootstyle=SUCCESS).grid(row=3, column=0, columnspan=3, pady=10, padx=20)

        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10, padx=20, sticky="ew")
        ttk.Button(button_frame, text="Clear", command=self.clear, bootstyle=WARNING).pack(side=LEFT, fill="x", expand=True, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.destroy, bootstyle=DANGER).pack(side=LEFT, fill="x", expand=True, padx=5)

    def upload_files(self):
        files = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel files", "*.xlsx *.xls")])
        if files:
            self.file_paths = list(files)
            self.files_listbox.delete(0, "end")
            for file in files:
                self.files_listbox.insert("end", os.path.basename(file))

    def combine_files(self):
        if not self.file_paths:
            messagebox.showwarning("No Files Selected", "Please upload files before proceeding.")
            return

        combined_data = []
        for file_path in self.file_paths:
            try:
                df = pd.read_excel(file_path)
                combined_data.append(df)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while reading {os.path.basename(file_path)}: {e}")
                return

        if combined_data:
            combined_df = pd.concat(combined_data)
            combined_df.sort_values(by="RegNum", inplace=True)
            combined_df.drop(columns=["Email"], inplace=True, errors="ignore")
            output_file = "Combined_File.xlsx"
            try:
                combined_df.to_excel(output_file, index=False)
                messagebox.showinfo("Success", f"Files combined successfully into {output_file}.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while saving the combined file: {e}")

    def clear(self):
        self.file_paths = []
        self.files_listbox.delete(0, "end")


if __name__ == "__main__":
    root = Tk()
    app = DashboardApp(root)
    root.mainloop()
