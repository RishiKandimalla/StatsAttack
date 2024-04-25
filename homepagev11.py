import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import csv
import win32com.client
import csv
from datetime import datetime
from tkinter import simpledialog
form_submissions = []
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
    def __init__(self, parent, show_home, data=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.pack(fill="both", expand=True)
        self.show_home = show_home
        self.create_form(data)

    def create_form(self, data=None):
        # Create a container for the scrollable frame and the buttons
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        # Create the scrollable frame as before
        scrollable_frame = ScrollableFrame(container)
        scrollable_frame.pack(fill="both", expand=True)

        self.entries = {}
        field_names = [
            "Entered By", "Primary User Name", "Primary User's Computing ID", "Date of the interaction",
            "Additional Staff", "Additional Users", "ARL Interaction Type",
            "Attendee Type", "Start Date", "Department",
            "Description", "Grant Related?", "Medium", "Pre-post-time",
            "Internal Notes", "Additional Notes", "RDS+SNE Group",
            "Referral", "School", "Session Duration", "Source/Software",
            "Staff", "Topic"
        ]

        for i, field in enumerate(field_names):
            label = ttk.Label(scrollable_frame.scrollable_frame, text=f"{field}:")
            label.grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            entry = ttk.Entry(scrollable_frame.scrollable_frame)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=10, pady=5)
            self.entries[field] = entry
            if data and field in data:
                entry.insert(0, data[field])

        # Place the buttons outside the scrollable frame, directly in the container
        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)

        submit_button = ttk.Button(button_frame, text="Submit", command=lambda: self.submit_form(self.entries, data))
        submit_button.pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame, text="Back to Home", command=self.show_home).pack(side=tk.RIGHT, padx=10)
        
    def submit_form(self, entries, data=None):
        form_data = {field: entry.get() for field, entry in entries.items()}
        if data:
            # Update the existing entry in form_submissions
            index = form_submissions.index(data)
            form_submissions[index] = form_data
        else:
            # Add a new entry to form_submissions
            form_submissions.append(form_data)
        message = "\n".join(f"{field}: {value}" for field, value in form_data.items())
        messagebox.showinfo("Form Submitted", message)
        self.show_home()  # Return to the home screen after submission

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
        ttk.Button(self, text="Import CSV", command=self.import_calendar_events).pack(side="top", anchor="ne", padx=10)
        # Table
        self.tree = ttk.Treeview(self, columns=("Date", "Name"), show="headings")
        self.tree.heading("Date", text="Entered By")
        self.tree.heading("Name", text="Date of Interaction")
        self.tree.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        # Back to Home Button
        ttk.Button(self, text="Back to Home", command=self.show_home).pack(side="bottom", pady=10)

    
    def fetch_outlook_events(self, start_date, end_date):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)  # 9 corresponds to the Calendar folder
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = "True"
        items = items.Restrict("[Start] >= '{}' AND [End] <= '{}'".format(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')))

        events = []
        for item in items:
            events.append({
                'Subject': item.Subject,
                'Start': item.Start.Format('%Y-%m-%d %H:%M'),
                'End': item.End.Format('%Y-%m-%d %H:%M'),
                'Organizer': item.Organizer,
                'Location': item.Location
            })
        return events
                
    def import_calendar_events(self):
        # Prompt user for date range
        start_date_str = simpledialog.askstring("Input", "Enter start date (YYYY-MM-DD):", parent=self)
        end_date_str = simpledialog.askstring("Input", "Enter end date (YYYY-MM-DD):", parent=self)
        
        try:
            # Parse the date strings to datetime objects
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
        except ValueError as e:
            messagebox.showerror("Date Error", "Invalid date format. Please use YYYY-MM-DD.")
            return
        
        # Fetch events from Outlook
        events = self.fetch_outlook_events(start_date, end_date)
        
        # Clear the treeview
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        # Insert the events into the treeview
        for event in events:
            self.tree.insert("", "end", values=(event["Subject"], event["Start"]))
            form_submissions.append(event);
        
                        
class DataScreen(ttk.Frame):
    def __init__(self, parent, show_home, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.show_home = show_home
        self.pack(fill="both", expand=True)
        self.create_data_screen()

    def create_data_screen(self):
        self.tree = ttk.Treeview(self, columns=("Primary User", "Date of Interaction"), show="headings")
        self.tree.heading("Primary User", text="Primary User")
        self.tree.heading("Date of Interaction", text="Date of Interaction")
        self.tree.pack(side="top", fill="both", expand=True, padx=10, pady=10)
        
        for submission in form_submissions:
            primary_user = submission.get("Primary User Name", "N/A")
            date_of_interaction = submission.get("Date of the interaction", "N/A")
            self.tree.insert("", "end", values=(primary_user, date_of_interaction), tags=("clickable",))
        self.tree.tag_bind("clickable", "<1>", self.on_item_click)
        
        #Add Export Button
        export_button = ttk.Button(self, text="Export to CSV", command=self.export_to_csv)
        export_button.pack(side="bottom", pady=10)
        
        # Add Delete Button
        delete_button = ttk.Button(self, text="Delete", command=self.delete_selected)
        delete_button.pack(side="bottom", pady=10)

        ttk.Button(self, text="Back to Home", command=self.show_home).pack(side="bottom", pady=10)

    def export_to_csv(self):
        # Ask user for location and name of the csv file to save
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return  # User cancelled; exit the function
        
        with open(file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            # Write the headers based on the form fields
            headers = [
                "Primary User Name", "Date of the interaction", "Additional Staff", "Additional Users",
                "ARL Interaction Type", "Attendee Type", "Start Date", "Department", "Description",
                "Grant Related?", "Medium", "Pre-post-time", "Internal Notes", "Additional Notes",
                "RDS+SNE Group", "Referral", "School", "Session Duration", "Source/Software", "Staff", "Topic"
            ]
            writer.writerow(headers)
            # Write data rows
            for submission in form_submissions:
                writer.writerow([submission.get(header, "") for header in headers])
            messagebox.showinfo("Export Successful", f"Data exported successfully to {file_path}")

    def on_item_click(self, event):
        item = self.tree.selection()[0]
        item_values = self.tree.item(item, "values")
        for submission in form_submissions:
            if submission["Primary User Name"] == item_values[0] and submission["Date of the interaction"] == item_values[1]:
                self.edit_submission(submission)
                break
            
    def delete_selected(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item, "values")
            # Find and remove the corresponding entry from form_submissions
            form_submissions[:] = [sub for sub in form_submissions if not (sub.get("Primary User Name", "N/A") == item_values[0] and sub.get("Date of the interaction", "N/A") == item_values[1])]
            # Remove the item from the Treeview
            self.tree.delete(selected_item)
    
    
    def edit_submission(self, data):
        # Destroy the current widgets and open the MainForm with the data for editing
        for widget in self.master.winfo_children():
            widget.destroy()
        MainForm(self.master, self.show_home, data)

class HomeScreen(tk.Frame):
    def __init__(self, parent, show_form, show_import, show_data, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.pack(fill="both", expand=True)
        
        # Set the button style options
        button_font = ("Arial", 10)  # Larger font size
        button_width = 5  # Width of the buttons
        button_height = 2  # Height of the buttons
        button_padx = 30  # Horizontal padding
        button_pady = 20  # Vertical padding

        # Create and pack the buttons with the specified styles and padding
        ttk.Button(self, text="Form", command=show_form, width=button_width, style='my.TButton').pack(side="left", padx=button_padx, pady=button_pady)
        ttk.Button(self, text="Data", command=show_data, width=button_width, style='my.TButton').pack(side="left", padx=button_padx, pady=button_pady)
        ttk.Button(self, text="Import", command=show_import, width=button_width, style='my.TButton').pack(side="left", padx=button_padx, pady=button_pady)

        # Configure the style for the buttons
        style = ttk.Style()
        style.configure('my.TButton', font=button_font, padding=[button_padx, button_pady])

def main():
    root = tk.Tk()
    root.title("Application")
    root.geometry("500x400")

    def show_home():
        for widget in root.winfo_children():
            widget.destroy()
        HomeScreen(root, lambda: show_form(), show_import, show_data)

    def show_form(data=None):
        for widget in root.winfo_children():
            widget.destroy()
        MainForm(root, show_home)

    def show_import():
        for widget in root.winfo_children():
            widget.destroy()
        ImportScreen(root, show_home)
    def show_data():
        for widget in root.winfo_children():
            widget.destroy()
        DataScreen(root, show_home)

    show_home()

    root.mainloop()

if __name__ == "__main__":
    main()