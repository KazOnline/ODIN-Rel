import tkinter as tk
from tkinter import messagebox
import sys
import os
import subprocess
from tkinter import ttk
from tkinter import filedialog
from pathlib import Path
import json
import shutil
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from tkinter import colorchooser
import win32print
import win32ui
from PIL import Image, ImageWin

#Core Assignments
SETTINGS = Path(__file__).with_name("settings.json")
with open(SETTINGS, 'r') as f:
    SETTINGS_DATA = json.load(f)
icon_dir = Path(__file__).parent / "icons"
ICON_PATH = icon_dir / "icon.png"

#Main Window
root = tk.Tk()
root.title(f"{SETTINGS_DATA.get('receiver', {}).get('app_name', ' ')} - Version {SETTINGS_DATA.get('receiver', {}).get('version', 'Version Unknown')}")
icon = tk.PhotoImage(file=str(ICON_PATH))
root.iconphoto(False, icon)
root.geometry("1200x600")
root.update_idletasks()

# Try to open maximized (cross-platform fallbacks)
try:
    root.state('zoomed')
except Exception:
    try:
        root.attributes('-zoomed', True)
    except Exception:
        pass

style = ttk.Style(root)
try:
    style.theme_use('default')
except Exception:
    try:
        style.theme_use(style.theme_names()[0])
    except Exception:
        pass

# Menu bar
menubar = tk.Menu(root)
root.config(menu=menubar)

# Menus
file_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="File", menu=file_menu)
edit_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Edit", menu=edit_menu)
about_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="About", menu=about_menu)

#Functions

def reselect_engineer():
    eng_select = tk.Toplevel(root)
    eng_select.title("Engineer Selection")
    eng_select.geometry("400x200")

    ttk.Label(eng_select, text="Select Engineer:").pack(pady=10)
    engineer_var = tk.StringVar()
    engineer_combo = ttk.Combobox(eng_select, textvariable=engineer_var, state="readonly")
    engineer_combo['values'] = [eng['name'] for eng in SETTINGS_DATA.get('engineers', [])]
    engineer_combo.pack(pady=10)

    def select_engineer():
        global selected_engineer
        selected_engineer = engineer_var.get()
        if selected_engineer:
            active_engineer.config(text=f"Active Engineer: {selected_engineer}")
            _show_list()
            eng_select.destroy()
        else:
            messagebox.showwarning("Selection Error", "Please select an engineer.")

    select_button = ttk.Button(eng_select, text="Select", command=select_engineer)
    select_button.pack(pady=10)

def _show_changelog():
    """Display changelog from changes.json in a popup with list and details view."""
    changes_path = Path(__file__).parent / "changes.json"
    
    if not changes_path.exists():
        messagebox.showwarning("No Changelog", "changes.json not found.")
        return
    
    try:
        with open(changes_path, 'r') as f:
            changelog_data = json.load(f)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load changes.json: {e}")
        return
    
    if not changelog_data:
        messagebox.showinfo("Empty", "No changelog entries found.")
        return
    
    # Create changelog window
    changelog_window = tk.Toplevel(root)
    changelog_window.title("Change Log")
    changelog_window.geometry("1000x600")
    
    # Main container with two panes
    main_frame = ttk.Frame(changelog_window)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Left pane - List of entries
    left_frame = ttk.Frame(main_frame)
    left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
    
    ttk.Label(left_frame, text="Versions", font=('TkDefaultFont', 10, 'bold')).pack(pady=(0, 5))
    
    list_canvas = tk.Canvas(left_frame)
    list_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=list_canvas.yview)
    list_container = ttk.Frame(list_canvas)
    
    list_container.bind("<Configure>", lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all")))
    
    list_canvas.create_window((0, 0), window=list_container, anchor="nw", width=list_canvas.winfo_reqwidth())
    list_canvas.configure(yscrollcommand=list_scrollbar.set)
    
    # Update canvas window width when canvas is resized
    def on_canvas_configure(event):
        list_canvas.itemconfig(list_canvas.find_withtag("all")[0], width=event.width)
    
    list_canvas.bind('<Configure>', on_canvas_configure)
    
    list_canvas.pack(side="left", fill="both", expand=True)
    list_scrollbar.pack(side="right", fill="y")
    
    # Right pane - Details
    right_frame = ttk.Frame(main_frame)
    right_frame.pack(side="right", fill="both", expand=True)
    
    ttk.Label(right_frame, text="Details", font=('TkDefaultFont', 10, 'bold')).pack(pady=(0, 5))
    
    details_canvas = tk.Canvas(right_frame, relief='solid', borderwidth=1)
    details_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=details_canvas.yview)
    details_container = ttk.Frame(details_canvas)
    
    details_container.bind("<Configure>", lambda e: details_canvas.configure(scrollregion=details_canvas.bbox("all")))
    
    details_canvas_window = details_canvas.create_window((0, 0), window=details_container, anchor="nw")
    details_canvas.configure(yscrollcommand=details_scrollbar.set)
    
    # Update details canvas window width when canvas is resized
    def on_details_configure(event):
        details_canvas.itemconfig(details_canvas_window, width=event.width)
    
    details_canvas.bind('<Configure>', on_details_configure)
    
    # Enable mousewheel scrolling for details canvas
    def on_mousewheel(event):
        details_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    details_canvas.bind_all("<MouseWheel>", on_mousewheel)
    changelog_window.protocol("WM_DELETE_WINDOW", lambda: [details_canvas.unbind_all("<MouseWheel>"), changelog_window.destroy()])
    
    details_canvas.pack(side="left", fill="both", expand=True)
    details_scrollbar.pack(side="right", fill="y")
    
    # Variable to track selected entry
    selected_entry = tk.StringVar()
    
    def show_details(entry):
        """Display details of selected changelog entry."""
        # Clear existing details
        for widget in details_container.winfo_children():
            widget.destroy()
        
        # Display entry details
        ttk.Label(details_container, text=f"Version: {entry.get('Version Number', 'N/A')}", font=('TkDefaultFont', 12, 'bold')).pack(anchor="w", fill="x", padx=10, pady=(10, 5))
        ttk.Label(details_container, text=f"Date: {entry.get('Date', 'N/A')}", font=('TkDefaultFont', 10)).pack(anchor="w", fill="x", padx=10, pady=(0, 5))
        ttk.Label(details_container, text=f"Title: {entry.get('Title', 'N/A')}", font=('TkDefaultFont', 10, 'italic')).pack(anchor="w", fill="x", padx=10, pady=(0, 10))
        
        ttk.Separator(details_container, orient='horizontal').pack(fill='x', padx=10, pady=10)
        
        ttk.Label(details_container, text="Changes:", font=('TkDefaultFont', 10, 'bold')).pack(anchor="w", fill="x", padx=10, pady=(0, 5))
        
        changes_text = entry.get('Changes', '')
        if changes_text:
            # Split by newlines and display each line
            lines = changes_text.split('\n')
            for line in lines:
                line = line.strip()
                if line:
                    if line.startswith('-'):
                        # Bullet point
                        change_frame = ttk.Frame(details_container)
                        change_frame.pack(fill="x", padx=20, pady=2)
                        
                        ttk.Label(change_frame, text="•", font=('TkDefaultFont', 10)).pack(side="left", padx=(0, 5))
                        ttk.Label(change_frame, text=line[1:].strip(), font=('TkDefaultFont', 9), wraplength=400, justify="left").pack(side="left", fill="x", expand=True)
                    else:
                        # Regular text
                        ttk.Label(details_container, text=line, font=('TkDefaultFont', 9), wraplength=450, justify="left").pack(anchor="w", fill="x", padx=20, pady=2)
        else:
            ttk.Label(details_container, text="No changes listed.", font=('TkDefaultFont', 9, 'italic')).pack(anchor="w", fill="x", padx=20, pady=5)
    
    # Populate list with entries
    for entry in changelog_data:
        entry_frame = ttk.Frame(list_container, relief='solid', borderwidth=1, padding=5)
        entry_frame.pack(fill="x", padx=5, pady=2)
        
        version = entry.get('Version Number', 'N/A')
        date = entry.get('Date', 'N/A')
        title = entry.get('Title', 'N/A')
        
        ttk.Label(entry_frame, text=f"{version}", font=('TkDefaultFont', 9, 'bold')).pack(anchor="w", fill="x")
        ttk.Label(entry_frame, text=f"{date}", font=('TkDefaultFont', 8)).pack(anchor="w", fill="x")
        ttk.Label(entry_frame, text=f"{title}", font=('TkDefaultFont', 8, 'italic')).pack(anchor="w", fill="x")
        
        # Make entry clickable
        def make_click_handler(e=entry):
            return lambda event: show_details(e)
        
        for widget in [entry_frame] + list(entry_frame.winfo_children()):
            widget.bind("<Button-1>", make_click_handler(entry))
            widget.configure(cursor="hand2")
    
    # Show first entry by default
    if changelog_data:
        show_details(changelog_data[0])
    
    # Close button
    ttk.Button(changelog_window, text="Close", command=changelog_window.destroy).pack(pady=(0, 10))

# Menus
file_menu.add_command(label="Exit", command=root.quit)
edit_menu.add_command(label="Change Active Engineer", command=reselect_engineer)
about_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", f"{SETTINGS_DATA.get('receiver', {}).get('app_name', ' ')}\nVersion {SETTINGS_DATA.get('receiver', {}).get('version', 'Version Unknown')}"))
about_menu.add_command(label="Changelog", command=_show_changelog)

reselect_engineer()

def _show_list():
    """Display database.json in a table format with selectable and colorable rows.""" 
    for widget in root.pack_slaves():
        if isinstance(widget, ttk.Frame) and widget != top_bar:
            widget.destroy()

    db_path = Path(__file__).parent / "database.json"
    
    if not db_path.exists():
        messagebox.showwarning("No Data", "database.json not found.")
        return
    
    try:
        with open(db_path, 'r') as f:
            data = json.load(f)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load database.json: {e}")
        return
    
    if not data:
        messagebox.showinfo("Empty", "No data in database.json")
        return
    
    # Load settings for colours and late_threshold
    list_colours = {'Priority': '#a85757', 'Late Dated': '#bd9582', 'Has Cert': '#57a857', 'Marked Complete': '#d3d3d3'}
    late_threshold = 7  # Default 7 days
    if SETTINGS.exists():
        try:
            with open(SETTINGS, 'r') as f:
                settings = json.load(f)
                list_colours = settings.get('list_colours', list_colours)
                late_threshold = settings.get('late_threshold', late_threshold)
        except Exception:
            pass
    
    # Sort data: Priority items first, then by Rec_Date (oldest first)
    def sort_key(entry):
        is_priority = entry.get('Priority', False)
        rec_date = entry.get('Rec_Date', '')
        # Convert priority to sort first (False=1, True=0)
        priority_sort = 0 if is_priority else 1
        return (priority_sort, rec_date)
    
    data.sort(key=sort_key)
    
    # Define columns to display
    headers = ['Job No', 'Cert_No', 'Cust_Ref', 'Serial_no', 'Manufacturer', 
               'Model_no', 'Description', 'Rec_Date', 'Customer', 'Status']
    
    # Create search frame at top
    search_frame = ttk.Frame(root)
    search_frame.pack(side="top", fill="x", padx=10, pady=(10, 0))
    
    # Dictionary to hold search entry variables
    search_vars = {}
    
    # Get unique list of assigned engineers from data
    assigned_engineers = set()
    for entry in data:
        assigned = entry.get('Assigned', '')
        if assigned:
            assigned_engineers.add(assigned)
    assigned_engineers = sorted(list(assigned_engineers))
    
    for idx, col in enumerate(headers):
        search_frame.grid_columnconfigure(idx, weight=1)
        
        ttk.Label(search_frame, text=col, font=('TkDefaultFont', 8)).grid(row=0, column=idx, sticky="ew", padx=2)
        
        search_var = tk.StringVar()
        search_vars[col] = search_var
        
        # Use dropdown for Assigned column
        if col == 'Assigned':
            combo = ttk.Combobox(search_frame, textvariable=search_var, values=[''] + assigned_engineers, state="readonly")
            combo.grid(row=1, column=idx, sticky="ew", padx=2, pady=2)
            combo.bind('<<ComboboxSelected>>', lambda e: filter_data())
        else:
            entry = ttk.Entry(search_frame, textvariable=search_var)
            entry.grid(row=1, column=idx, sticky="ew", padx=2, pady=2)
            entry.bind('<KeyRelease>', lambda e: filter_data())
    
    # Create frame for table
    table_frame = ttk.Frame(root)
    table_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)
    
    # Create Treeview with scrollbars
    tree_scroll_y = ttk.Scrollbar(table_frame, orient="vertical")
    tree_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal")
    
    tree = ttk.Treeview(
        table_frame,
        columns=headers,
        show="headings",
        yscrollcommand=tree_scroll_y.set,
        xscrollcommand=tree_scroll_x.set,
        selectmode="extended"
    )
    
    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)
    
    # Configure tags for different colours
    tree.tag_configure('priority', background=list_colours.get('Priority', '#ff0000'))
    tree.tag_configure('late_dated', background=list_colours.get('Late Dated', '#ffcc00'))
    tree.tag_configure('has_cert', background=list_colours.get('Has Cert', '#00ff00'))
    tree.tag_configure('marked_complete', background=list_colours.get('Marked Complete', '#d3d3d3'))
    
    # Configure column headings with initial width
    for col in headers:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="w")
    
    # Get today's date for comparison
    today = datetime.now().date()
    
    # Variable to track current count
    count_var = tk.StringVar()
    
    def filter_data():
        """Filter and populate tree based on search criteria."""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Get search filters
        filters = {col: search_vars[col].get().lower() for col in headers}
        
        # Filter and insert data
        visible_count = 0
        for entry in data:
            # Filter by active engineer first
            if selected_engineer != "None" and entry.get('Assigned', '') != selected_engineer:
                continue
            
            # Check if entry matches all filters
            match = True
            for col, search_text in filters.items():
                if search_text:
                    value = str(entry.get(col, "")).lower()
                    if search_text not in value:
                        match = False
                        break
            
            if not match:
                continue
            
            visible_count += 1
            values = []
            rec_date_obj = None
            
            for col in headers:
                value = entry.get(col, "")
                
                # Format Rec_Date as DD/MM/YYYY and store datetime object
                if col == 'Rec_Date' and value:
                    try:
                        if isinstance(value, str):
                            dt = datetime.fromisoformat(value)
                        else:
                            dt = value
                        rec_date_obj = dt.date()
                        value = dt.strftime('%d/%m/%Y')
                    except Exception:
                        pass
                
                # Format Last Update as DD/MM/YYYY HH:MM
                if col == 'Last Update' and value:
                    try:
                        if isinstance(value, str):
                            dt = datetime.fromisoformat(value)
                        else:
                            dt = value
                        value = dt.strftime('%d/%m/%Y %H:%M')
                    except Exception:
                        pass
                
                values.append(value)
            
            # Determine tag based on priority rules
            tag = ()
            is_priority = entry.get('Priority', False)
            cert_no = entry.get('Cert_No', '')
            is_late = rec_date_obj and (today - rec_date_obj).days > late_threshold
            is_marked_complete = entry.get('Status', '') == 'N'
            
            if is_marked_complete:
                tag = ('marked_complete',)
            elif is_priority:
                tag = ('priority',)
            elif is_late:
                tag = ('late_dated',)
            elif cert_no:
                tag = ('has_cert',)
            
            tree.insert("", "end", values=values, tags=tag)
        
        # Update count display
        count_var.set(f"Current instruments: {visible_count}")
        
        # Auto-size columns to fit content
        for col in headers:
            max_width = len(col) * 10  # Header width
            for item in tree.get_children():
                item_text = str(tree.set(item, col))
                text_width = len(item_text) * 8
                max_width = max(max_width, text_width)
            tree.column(col, width=min(max_width + 20, 400))  # Cap at 400px
    
    # Initial population
    filter_data()
    
    # Pack scrollbars and treeview
    tree_scroll_y.pack(side="right", fill="y")
    tree_scroll_x.pack(side="bottom", fill="x")
    tree.pack(side="left", fill="both", expand=True)
    
    # Store reference for later color manipulation
    root._current_tree = tree
    
    # Options bar at bottom
    options_frame = ttk.Frame(root, height=60)
    options_frame.pack(side="bottom", fill="x", padx=10, pady=10)
    options_frame.pack_propagate(False)
    
    # Load icons for options buttons
    options_icon_dir = Path(__file__).parent / "icons"
    options_icons = {
        "Mark Complete": "done.png",
        "Mark Incomplete": "undone.png",
        "Refresh List": "refresh.png"
    }
    
    # Keep image references
    if not hasattr(root, "_options_img_refs"):
        root._options_img_refs = []

    def mark_complete():
        selected_items = root._current_tree.selection()
        if not selected_items:
            messagebox.showinfo("No Selection", "Please select one or more entries to mark as complete.")
            return
        
        db_path = Path(__file__).parent / "database.json"
        with open(db_path, 'r') as f:
            data = json.load(f)
        
        for item in selected_items:
            values = root._current_tree.item(item, 'values')
            job_no = values[0]
            for entry in data:
                if entry.get('Job No') == job_no:
                    entry['Status'] = 'N'
                    entry['Last Update'] = datetime.now().isoformat()
                    break
        
        with open(db_path, 'w') as f:
            json.dump(data, f, indent=4)
        
        _show_list()    

    def mark_incomplete():
        selected_items = root._current_tree.selection()
        if not selected_items:
            messagebox.showinfo("No Selection", "Please select one or more entries to mark as complete.")
            return
        
        db_path = Path(__file__).parent / "database.json"
        with open(db_path, 'r') as f:
            data = json.load(f)
        
        for item in selected_items:
            values = root._current_tree.item(item, 'values')
            job_no = values[0]
            for entry in data:
                if entry.get('Job No') == job_no:
                    entry['Status'] = 'E'
                    entry['Last Update'] = datetime.now().isoformat()
                    break
        
        with open(db_path, 'w') as f:
            json.dump(data, f, indent=4)
        
        _show_list()    

    def refresh_list():
        _show_list()
    
    # Map button labels to commands
    button_commands = {
        "Mark Complete": mark_complete,
        "Mark Incomplete": mark_incomplete,
        "Refresh List": refresh_list
    }
    
    # Create buttons with icons
    for label, fname in options_icons.items():
        img = None
        fpath = options_icon_dir / fname
        try:
            if fpath.exists():
                img = tk.PhotoImage(file=str(fpath))
                root._options_img_refs.append(img)
        except Exception:
            img = None
        
        cmd = button_commands.get(label)
        
        if img:
            btn = ttk.Button(options_frame, text=label, image=img, compound="left", command=cmd)
        else:
            btn = ttk.Button(options_frame, text=label, command=cmd)
        
        btn.pack(side="left", padx=5)

top_bar = ttk.Frame(root, height=40)
top_bar.pack(side="top", fill="x")
top_bar.pack_propagate(False)

active_engineer = ttk.Label(top_bar, text="Active Engineer: None", font=('TkDefaultFont', 10, 'bold'))
active_engineer.pack(side="left", padx=10)

ttk.Separator(root, orient="horizontal").pack(side="top", fill="x", padx=5)

root.mainloop()