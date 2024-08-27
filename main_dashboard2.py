import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

class MenuDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Main Dashboard")
        self.root.geometry("400x250")
        self.root.resizable(False, False)

        # Create GUI components
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        ttk.Label(self.root, text="Main Dashboard",
                  font=("Helvetica", 18, "bold"), bootstyle="primary").pack(
            pady=20)

        # File Combiner Button
        ttk.Button(self.root, text="Open File Combiner",
                   command=self.open_file_combiner, bootstyle=INFO).pack(
            pady=10, fill=X, padx=20)

        # Exit Button
        ttk.Button(self.root, text="Exit", command=self.root.quit,
                   bootstyle=DANGER).pack(pady=10, fill=X, padx=20)

    def open_file_combiner(self):
        # Import the FileCombinerApp class
        from file_combiner_OLD import FileCombinerApp
        # Create a new window for FileCombinerApp
        combiner_window = ttk.Window(themename="flatly")
        FileCombinerApp(combiner_window)
        combiner_window.mainloop()

# Create the main window for the MenuDashboard
root = ttk.Window(themename="flatly")
app = MenuDashboard(root)
root.mainloop()
