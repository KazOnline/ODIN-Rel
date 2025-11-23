import tkinter as tk
from tkinter import ttk, messagebox
import json
from datetime import datetime
import os
from tkcalendar import DateEntry
from pathlib import Path

CHANGE_PATH = Path(__file__).with_name("changes.json")

class ChangeLogEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Change Log Editor")
        self.root.geometry("800x600")
        
        # Main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left side - Entry list
        list_frame = ttk.LabelFrame(main_frame, text="Existing Entries", padding="5")
        list_frame.grid(row=0, column=0, rowspan=5, sticky=(tk.N, tk.S, tk.W, tk.E), padx=(0, 10))
        
        # Listbox with scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.entry_listbox = tk.Listbox(list_frame, width=30, height=20, yscrollcommand=scrollbar.set)
        self.entry_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.entry_listbox.yview)
        
        self.entry_listbox.bind('<<ListboxSelect>>', self.on_entry_select)
        
        # Right side - Entry form
        form_frame = ttk.Frame(main_frame)
        form_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Date (date-picker)
        ttk.Label(form_frame, text="Date:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.date_picker = DateEntry(form_frame, width=37, background='darkblue',
                                      foreground='white', borderwidth=2,
                                      date_pattern='yyyy-mm-dd')
        self.date_picker.grid(row=0, column=1, pady=5)
        
        # Title
        ttk.Label(form_frame, text="Title:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.title_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.title_var, width=40).grid(row=1, column=1, pady=5)
        
        # Version Number
        ttk.Label(form_frame, text="Version Number:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.version_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.version_var, width=40).grid(row=2, column=1, pady=5)
        
        # Changes (Text)
        ttk.Label(form_frame, text="Changes:").grid(row=3, column=0, sticky=tk.NW, pady=5)
        
        # Text widget
        self.changes_text = tk.Text(form_frame, width=50, height=15, wrap=tk.WORD)
        self.changes_text.grid(row=3, column=1, pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Save New Entry", command=self.save_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Update Entry", command=self.update_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear", command=self.clear_fields).pack(side=tk.LEFT, padx=5)
        
        self.current_index = None
        self.load_entries()
    
    def load_entries(self):
        """Load and display all entries in the listbox"""
        self.entry_listbox.delete(0, tk.END)
        
        if os.path.exists(CHANGE_PATH):
            try:
                with open(CHANGE_PATH, "r") as f:
                    content = f.read().strip()
                    if content:
                        self.data = json.loads(content)
                        for entry in self.data:
                            display_text = f"{entry.get('Version Number', 'N/A')} - {entry.get('Title', 'Untitled')}"
                            self.entry_listbox.insert(tk.END, display_text)
                    else:
                        self.data = []
            except json.JSONDecodeError:
                self.data = []
        else:
            self.data = []
    
    def on_entry_select(self, event):
        """Load selected entry into the form"""
        selection = self.entry_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_index = index
            entry = self.data[index]
            
            # Populate fields
            self.title_var.set(entry.get("Title", ""))
            self.version_var.set(entry.get("Version Number", ""))
            self.changes_text.delete("1.0", tk.END)
            self.changes_text.insert("1.0", entry.get("Changes", ""))
            
            # Set date
            try:
                date_obj = datetime.strptime(entry.get("Date", ""), "%Y-%m-%d")
                self.date_picker.set_date(date_obj)
            except ValueError:
                self.date_picker.set_date(datetime.now())
    
    def clear_fields(self):
        """Clear all input fields"""
        self.title_var.set("")
        self.version_var.set("")
        self.changes_text.delete("1.0", tk.END)
        self.date_picker.set_date(datetime.now())
        self.current_index = None
        self.entry_listbox.selection_clear(0, tk.END)
    
    def save_entry(self):
        try:
            version = self.version_var.get().strip()
            title = self.title_var.get().strip()
            date = self.date_picker.get_date().strftime("%Y-%m-%d")
            changes = self.changes_text.get("1.0", tk.END).strip()
            
            if not version or not title or not changes:
                messagebox.showerror("Error", "Title, Version Number and Changes are required!")
                return
            
            entry = {
                "Date": date,
                "Title": title,
                "Version Number": version,
                "Changes": changes
            }
            
            # Add new entry
            self.data.append(entry)
            
            # Save to file
            with open(CHANGE_PATH, "w") as f:
                json.dump(self.data, f, indent=2)
            
            messagebox.showinfo("Success", "Entry saved successfully!")
            
            # Refresh list and clear fields
            self.load_entries()
            self.clear_fields()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save entry:\n{str(e)}")
    
    def update_entry(self):
        """Update the currently selected entry"""
        if self.current_index is None:
            messagebox.showwarning("Warning", "Please select an entry to update!")
            return
        
        try:
            version = self.version_var.get().strip()
            title = self.title_var.get().strip()
            date = self.date_picker.get_date().strftime("%Y-%m-%d")
            changes = self.changes_text.get("1.0", tk.END).strip()
            
            if not version or not title or not changes:
                messagebox.showerror("Error", "Title, Version Number and Changes are required!")
                return
            
            # Update entry
            self.data[self.current_index] = {
                "Date": date,
                "Title": title,
                "Version Number": version,
                "Changes": changes
            }
            
            # Save to file
            with open(CHANGE_PATH, "w") as f:
                json.dump(self.data, f, indent=2)
            
            messagebox.showinfo("Success", "Entry updated successfully!")
            
            # Refresh list and clear fields
            self.load_entries()
            self.clear_fields()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update entry:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ChangeLogEditor(root)
    root.mainloop()
