import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

    def _on_canvas_configure(self, event):
        canvas_width = event.width
        self.canvas.itemconfig("inner_frame", width=canvas_width)


class MainForm(ttk.Frame):
    def __init__(self, parent,show_home, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.pack(fill="both", expand=True)
        self.show_home = show_home
        self.create_form()

    def create_form(self):
        scrollable_frame = ScrollableFrame(self)
        scrollable_frame.pack(fill="both", expand=True)

        entries = {}
        field_names = [
            "Name", "Email", "Phone", "Date", "Computing ID", "Session Duration",
            "Additional Users", "Attendee Type"
        ]

        for i, field in enumerate(field_names):
            label = ttk.Label(scrollable_frame.scrollable_frame, text=f"{field}:")
            label.grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            entry = ttk.Entry(scrollable_frame.scrollable_frame)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=10, pady=5)
            entries[field] = entry

        submit_button = ttk.Button(scrollable_frame.scrollable_frame, text="Submit", command=lambda: self.submit_form(entries))
        submit_button.grid(row=len(field_names), column=0, columnspan=2, pady=10)
        ttk.Button(self, text="Back to Home", command=self.show_home).pack(side="bottom", pady=10)


    def submit_form(self, entries):
        form_data = {field: entry.get() for field, entry in entries.items()}
        message = "\n".join(f"{field}: {value}" for field, value in form_data.items())
        messagebox.showinfo("Form Submitted", message)


class ImportScreen(ttk.Frame):
    def __init__(self, parent, show_home, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.show_home = show_home
        self.pack(fill="both", expand=True)
        self.create_import_screen()

    def create_import_screen(self):
        # Title and existing code remains unchanged
        ttk.Label(self, text="Calendar Events").pack(side="top", fill="x", pady=10)

        # Import Button
        ttk.Button(self, text="Import CSV", command=self.import_csv).pack(side="top", anchor="ne", padx=10)
        # Table
        self.tree = ttk.Treeview(self, columns=("Date", "Name"), show="headings")
        self.tree.heading("Date", text="Date of the interaction")
        self.tree.heading("Name", text="Entered By")
        self.tree.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        # Back to Home Button
        ttk.Button(self, text="Back to Home", command=self.show_home).pack(side="bottom", pady=10)

    def import_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                df = pd.read_csv(file_path)
                # Clear the treeview
                for i in self.tree.get_children():
                    self.tree.delete(i)
                # Inserting new rows
                for _, row in df.iterrows():
                    # Extract data from the specified columns
                    date_of_interaction = row.get("Date of the interaction", "N/A")
                    entered_by = row.get("Entered By", "N/A")
                    self.tree.insert("", "end", values=(date_of_interaction, entered_by))
            except Exception as e:
                messagebox.showerror("Import Error", str(e))


class HomeScreen(ttk.Frame):
    def __init__(self, parent, show_form, show_import, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.pack(fill="both", expand=True)
        ttk.Button(self, text="Form", command=show_form).pack(side="left", padx=20, pady=20)
        ttk.Button(self, text="Data").pack(side="left", padx=20, pady=20)
        ttk.Button(self, text="Import", command=show_import).pack(side="left", padx=20, pady=20)

def main():
    root = tk.Tk()
    root.title("Application")
    root.geometry("500x400")

    def show_home():
        for widget in root.winfo_children():
            widget.destroy()
        HomeScreen(root, show_form, show_import)

    def show_form():
        for widget in root.winfo_children():
            widget.destroy()
        MainForm(root, show_home)

    def show_import():
        for widget in root.winfo_children():
            widget.destroy()
        ImportScreen(root, show_home)

    show_home()

    root.mainloop()

if __name__ == "__main__":
    main()