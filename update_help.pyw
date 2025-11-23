import json
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
from pathlib import Path

HELP_PATH = Path(__file__).with_name("help.json")

class HelpEntryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Help Entry Manager")
        self.root.geometry("1200x600")
        self.edit_index = None
        
        # Left frame for form
        form_frame = tk.Frame(root)
        form_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Date
        tk.Label(form_frame, text="Date:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.date_entry = DateEntry(form_frame, width=20, background='darkblue',
                                    foreground='white', borderwidth=2,
                                    date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        # Title
        tk.Label(form_frame, text="Title:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.title_entry = tk.Entry(form_frame, width=50)
        self.title_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        
        # Content
        tk.Label(form_frame, text="Content:").grid(row=2, column=0, sticky="nw", padx=10, pady=5)
        self.content_text = tk.Text(form_frame, width=50, height=10, wrap=tk.WORD)
        self.content_text.grid(row=2, column=1, sticky="nsew", padx=10, pady=5)
        
        # Scrollbar for content
        scrollbar = ttk.Scrollbar(form_frame, command=self.content_text.yview)
        scrollbar.grid(row=2, column=2, sticky="ns", pady=5)
        self.content_text.config(yscrollcommand=scrollbar.set)
        
        # Buttons
        button_frame = tk.Frame(form_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.save_button = tk.Button(button_frame, text="Add Entry", command=self.add_entry, 
                 bg="green", fg="white", width=15)
        self.save_button.pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Clear", command=self.clear_fields,
                 width=15).pack(side=tk.LEFT, padx=5)
        
        # Right frame for list
        list_frame = tk.Frame(root)
        list_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        tk.Label(list_frame, text="Existing Entries:", font=("Arial", 10, "bold")).pack(anchor="w")
        
        # Listbox with scrollbar
        list_scroll = ttk.Scrollbar(list_frame)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.entry_listbox = tk.Listbox(list_frame, yscrollcommand=list_scroll.set, width=40)
        self.entry_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.config(command=self.entry_listbox.yview)
        
        # List buttons
        list_button_frame = tk.Frame(list_frame)
        list_button_frame.pack(pady=5)
        
        tk.Button(list_button_frame, text="Edit", command=self.edit_entry, 
                 bg="blue", fg="white", width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(list_button_frame, text="Delete", command=self.delete_entry,
                 bg="red", fg="white", width=12).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        root.grid_columnconfigure(0, weight=2)
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=1)
        form_frame.grid_rowconfigure(2, weight=1)
        
        # Load entries
        self.load_entries()
    
    def load_entries(self):
        self.entry_listbox.delete(0, tk.END)
        try:
            with open(HELP_PATH, "r") as f:
                content = f.read().strip()
                if content:
                    self.data = json.loads(content)
                else:
                    self.data = []
        except (FileNotFoundError, json.JSONDecodeError):
            self.data = []
        
        for entry in self.data:
            self.entry_listbox.insert(tk.END, f"{entry['date']} - {entry['title']}")
    
    def add_entry(self):
        date = self.date_entry.get_date().strftime('%Y-%m-%d')
        title = self.title_entry.get().strip()
        content = self.content_text.get("1.0", tk.END).strip()
        
        if not title or not content:
            messagebox.showwarning("Missing Information", "Please fill in all fields.")
            return
        
        entry = {
            "date": date,
            "title": title,
            "content": content
        }
        
        try:
            if self.edit_index is not None:
                # Update existing entry
                self.data[self.edit_index] = entry
                self.edit_index = None
                self.save_button.config(text="Add Entry", bg="green")
            else:
                # Add new entry
                self.data.append(entry)
            
            # Save to file
            with open(HELP_PATH, "w") as f:
                json.dump(self.data, f, indent=4)
            
            messagebox.showinfo("Success", "Help entry saved successfully!")
            self.clear_fields()
            self.load_entries()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save entry: {str(e)}")
    
    def edit_entry(self):
        selection = self.entry_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an entry to edit.")
            return
        
        index = selection[0]
        self.edit_index = index
        entry = self.data[index]
        
        # Populate fields
        self.date_entry.set_date(datetime.strptime(entry['date'], '%Y-%m-%d'))
        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, entry['title'])
        self.content_text.delete("1.0", tk.END)
        self.content_text.insert("1.0", entry['content'])
        
        self.save_button.config(text="Update Entry", bg="orange")
    
    def delete_entry(self):
        selection = self.entry_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an entry to delete.")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this entry?"):
            index = selection[0]
            self.data.pop(index)
            
            try:
                with open(HELP_PATH, "w") as f:
                    json.dump(self.data, f, indent=4)
                messagebox.showinfo("Success", "Entry deleted successfully!")
                self.load_entries()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete entry: {str(e)}")
    
    def clear_fields(self):
        self.date_entry.set_date(datetime.now())
        self.title_entry.delete(0, tk.END)
        self.content_text.delete("1.0", tk.END)
        self.edit_index = None
        self.save_button.config(text="Add Entry", bg="green")


if __name__ == "__main__":
    root = tk.Tk()
    app = HelpEntryGUI(root)
    root.mainloop()