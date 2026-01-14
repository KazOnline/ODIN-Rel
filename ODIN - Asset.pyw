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
REG_PATH = Path(__file__).parent / "assets" / "registry.json"
INDEX_PATH = Path(__file__).parent / "assets" / "index.json"
TIC_PATH = Path(__file__).parent / "assets" / "tickets.json"
COST_PATH = Path(__file__).parent / "assets" / "costs.json"
ASSET_SETTINGS_PATH = Path(__file__).parent / "assets" / "asset_settings.json"
SUPPLIER_LIST_PATH = Path(__file__).parent / "assets" / "suppliers.json"

# Load asset settings
with open(ASSET_SETTINGS_PATH, 'r') as f:
    ASSET_SETTINGS = json.load(f)

# Create backup directory if it doesn't exist
BACKUP_DIR = Path(__file__).parent / "history" / "backups"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
MAX_BACKUPS = 10

# Ticket constants
TICKET_TYPES = ["Calibration", "Repair", "General", "Relocate"]
ACTION_TYPES = ["Comment", "File Attach", "Calibration Certificate", "Dispatch", "Return", "Registry Update"]

# Initialize data files if they don't exist
if not REG_PATH.exists():
    with open(REG_PATH, 'w') as f:
        json.dump([], f)

if not INDEX_PATH.exists():
    with open(INDEX_PATH, 'w') as f:
        json.dump([], f)

if not TIC_PATH.exists():
    with open(TIC_PATH, 'w') as f:
        json.dump({"next_ticket_number": 1, "tickets": []}, f, indent=2)

if not COST_PATH.exists():
    with open(COST_PATH, 'w') as f:
        json.dump({"next_cost_number": 1, "costs": []}, f, indent=2)

if not SUPPLIER_LIST_PATH.exists():
    with open(SUPPLIER_LIST_PATH, 'w') as f:
        json.dump({"suppliers": []}, f, indent=2)

# Backup management function
def create_backup(file_path):
    """Create a timestamped backup of a file and maintain max 10 backups"""
    if not file_path.exists():
        return
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
    backup_path = BACKUP_DIR / backup_name
    
    # Create the backup
    shutil.copy2(file_path, backup_path)
    
    # Get all backups for this file type
    pattern = f"{file_path.stem}_*{file_path.suffix}"
    backups = sorted(BACKUP_DIR.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    
    # Remove old backups if we exceed MAX_BACKUPS
    for old_backup in backups[MAX_BACKUPS:]:
        old_backup.unlink()

def save_json(file_path, data):
    """Save JSON data to file with automatic backup"""
    create_backup(file_path)
    with open(file_path, 'w') as f:
        json.dump(data, f, indent=2)

# Helper function to manage ticket folders and copy files
def create_ticket_folder(ticket_id):
    """Create folder for ticket if it doesn't exist"""
    ticket_dir = ASSET_SETTINGS.get('ticket_files_directory', str(Path(__file__).parent / "tickets"))
    ticket_folder = Path(ticket_dir) / ticket_id
    ticket_folder.mkdir(parents=True, exist_ok=True)
    return ticket_folder

def copy_file_to_ticket(file_path, ticket_id):
    """Copy a file to the ticket folder and return the new path"""
    if not file_path or not Path(file_path).exists():
        return None
    
    ticket_folder = create_ticket_folder(ticket_id)
    source = Path(file_path)
    destination = ticket_folder / source.name
    
    # Handle duplicate filenames
    counter = 1
    while destination.exists():
        destination = ticket_folder / f"{source.stem}_{counter}{source.suffix}"
        counter += 1
    
    shutil.copy2(source, destination)
    return str(destination)

load = tk.Tk()
load.title("ODIN Asset - Loading...")
load.geometry("600x600")
icon = tk.PhotoImage(file=str(ICON_PATH))
load.iconphoto(False, icon)
logo = tk.PhotoImage(file=str(icon_dir / "logo.png"))
logo_label = ttk.Label(load, image=logo)
logo_label.pack(pady=10)
ttk.Label(load, text="ODIN - Asset is loading, please wait...", font=('TkDefaultFont', 12)).pack(expand=True)
load_time = 2000  # milliseconds
load.update()
load.after(load_time, load.destroy)
load.mainloop()

#Main Window

class AssetManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{SETTINGS_DATA.get('asset', {}).get('app_name', 'ODIN Asset')} - Version {SETTINGS_DATA.get('asset', {}).get('version', 'Version Unknown')}")
        icon = tk.PhotoImage(file=str(ICON_PATH))
        self.root.iconphoto(False, icon)
        self.root.geometry("1400x900")
        
        # Try to open maximized (cross-platform fallbacks)
        try:
            self.root.state('zoomed')
        except Exception:
            try:
                self.root.attributes('-zoomed', True)
            except Exception:
                pass


        # Current user
        self.current_user = tk.StringVar()
        self.current_user.set(ASSET_SETTINGS.get('users', ['User 1'])[0])
        
        # Create menu bar
        self.create_menu()
        
        # Create main frame with user selector
        self.create_header()
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create tabs
        self.create_assets_tab()
        self.create_tickets_tab()
        self.create_expenditure_tab()
        
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Add Asset", command=self.add_new_asset)
        file_menu.add_separator()
        file_menu.add_command(label="Reload Data", command=self.reload_data)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Edit Asset Settings", command=self.edit_asset_settings)
        settings_menu.add_command(label="View/Edit Suppliers", command=self.supplier_window)

    def supplier_window(self):
        SupplierWindow(self.root, None)

    def create_header(self):
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(header_frame, text="Current User:", font=('TkDefaultFont', 10, 'bold')).pack(side='left', padx=5)
        user_dropdown = ttk.Combobox(header_frame, textvariable=self.current_user, 
                                     values=ASSET_SETTINGS.get('users', ['User 1']), 
                                     state='readonly', width=20)
        user_dropdown.pack(side='left', padx=5)
        
    def create_assets_tab(self):
        assets_frame = ttk.Frame(self.notebook)
        self.notebook.add(assets_frame, text="Assets")
        
        # Initialize search vars
        self.assets_search_vars = {}
        
        # Toolbar
        toolbar = ttk.Frame(assets_frame)
        toolbar.pack(fill='x', padx=5, pady=5)
        ttk.Button(toolbar, text="Refresh", command=self.load_assets).pack(side='left', padx=2)
        ttk.Button(toolbar, text="Edit Registry", command=self.edit_registry_item).pack(side='left', padx=2)      
        # Info label
        ttk.Label(toolbar, text="(Index fields are read-only, Registry fields are editable)", 
                 font=('TkDefaultFont', 8, 'italic')).pack(side='left', padx=20)
        
        # Search frame (will be populated after columns are known)
        self.assets_search_frame = ttk.Frame(assets_frame)
        self.assets_search_frame.pack(fill='x', padx=5, pady=5)
        
        # Treeview
        tree_frame = ttk.Frame(assets_frame)
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient='vertical')
        scrollbar_x = ttk.Scrollbar(tree_frame, orient='horizontal')
        
        self.assets_tree = ttk.Treeview(tree_frame, yscrollcommand=scrollbar_y.set, 
                                        xscrollcommand=scrollbar_x.set)
        scrollbar_y.config(command=self.assets_tree.yview)
        scrollbar_x.config(command=self.assets_tree.xview)
        
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.assets_tree.pack(fill='both', expand=True)
        
        # Action bar at bottom
        action_bar = ttk.Frame(assets_frame)
        action_bar.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(action_bar, text="Asset Actions:", font=('TkDefaultFont', 10, 'bold')).pack(side='left', padx=5)
        ttk.Button(action_bar, text="New Ticket", command=self.new_ticket_from_asset).pack(side='left', padx=2)
        ttk.Button(action_bar, text="View Tickets", command=self.view_asset_tickets).pack(side='left', padx=2)
        ttk.Button(action_bar, text="Change Location", command=self.change_asset_location).pack(side='left', padx=2)
        ttk.Button(action_bar, text="Change Status", command=self.change_asset_status).pack(side='left', padx=2)
        ttk.Button(action_bar, text="Edit Registry", command=self.edit_registry_item).pack(side='left', padx=2)
        ttk.Button(action_bar, text="View Details", command=self.view_asset_details).pack(side='left', padx=2)
        
        self.load_assets()
        
    def create_tickets_tab(self):
        tickets_frame = ttk.Frame(self.notebook)
        self.notebook.add(tickets_frame, text="Tickets")
        
        # Split into list and detail
        paned = ttk.PanedWindow(tickets_frame, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Left: Ticket list
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        
        # Toolbar
        toolbar = ttk.Frame(left_frame)
        toolbar.pack(fill='x', padx=5, pady=5)
        ttk.Button(toolbar, text="New Ticket", command=self.create_new_ticket).pack(side='left', padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_tickets).pack(side='left', padx=2)
        ttk.Button(toolbar, text="Edit Ticket", command=self.edit_selected_ticket).pack(side='left', padx=2)

        # Search frame for column-specific search
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(fill='x', padx=5, pady=5)
        
        # Dictionary to hold search vars
        self.ticket_search_vars = {}
        
        # Ticket columns
        ticket_columns = ['ticket_id', 'title','type', 'asset', 'date', 'user', 'status', 'po_number']
        
        for idx, col in enumerate(ticket_columns):
            search_frame.grid_columnconfigure(idx, weight=1)
            
            ttk.Label(search_frame, text=col.replace('_', ' ').title(), 
                     font=('TkDefaultFont', 8)).grid(row=0, column=idx, sticky='ew', padx=2)
            
            search_var = tk.StringVar()
            self.ticket_search_vars[col] = search_var
            
            entry = ttk.Entry(search_frame, textvariable=search_var, width=10)
            entry.grid(row=1, column=idx, sticky='ew', padx=2, pady=2)
            entry.bind('<KeyRelease>', lambda e: self.filter_tickets())
        
        # Ticket list
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        self.ticket_tree = ttk.Treeview(tree_frame, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.ticket_tree.yview)
        
        scrollbar.pack(side='right', fill='y')
        self.ticket_tree.pack(fill='both', expand=True)
        
        self.ticket_tree['columns'] = ('ticket_id', 'type', 'asset', 'date', 'user', 'status', 'po_number')
        self.ticket_tree.column('#0', width=0, stretch=False)
        
        for col in self.ticket_tree['columns']:
            self.ticket_tree.column(col, width=100)
            self.ticket_tree.heading(col, text=col.replace('_', ' ').title())
        
        self.ticket_tree.bind('<<TreeviewSelect>>', self.on_ticket_select)
        
        # Right: Ticket detail
        self.detail_frame = ttk.Frame(paned)
        paned.add(self.detail_frame, weight=2)
        
        ttk.Label(self.detail_frame, text="Select a ticket to view details", 
                 font=('TkDefaultFont', 12)).pack(pady=20)
        
        self.load_tickets()
        
    def create_expenditure_tab(self):
        expenditure_frame = ttk.Frame(self.notebook)
        self.notebook.add(expenditure_frame, text="Expenditure")
        
        #Split into a paned window
        paned = ttk.Panedwindow(expenditure_frame, orient='vertical')
        paned.pack(fill='both', expand=True, padx=5, pady=5)

        ttk.Button(expenditure_frame, text="Refresh Summary", command=self.load_capex_summary).pack(pady=2, padx=5, anchor='e')

        #Overview Table
        overview_frame = ttk.Frame(paned)
        paned.add(overview_frame, weight=1)
        ttk.Label(overview_frame, text="Expenditure Summary", font=('TkDefaultFont', 12, 'bold')).pack(pady=5)
        overview_tree_frame = ttk.Frame(overview_frame)
        overview_tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar_y = ttk.Scrollbar(overview_tree_frame, orient='vertical')
        scrollbar_x = ttk.Scrollbar(overview_tree_frame, orient='horizontal')
        self.overview_tree = ttk.Treeview(overview_tree_frame, yscrollcommand=scrollbar_y.set,
                                        xscrollcommand=scrollbar_x.set)
        scrollbar_y.config(command=self.overview_tree.yview)
        scrollbar_x.config(command=self.overview_tree.xview)
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.overview_tree.pack(fill='both', expand=True)

        self.overview_tree['columns'] = ('Code', 'Name', 'Total Budget', 'Number of Costs', 'Remaining Budget', 'Total Used', 'Percentage Used')

        self.overview_tree.column('#0', width=0, stretch=False)
        for col in self.overview_tree['columns']:
            self.overview_tree.column(col, width=150)
            self.overview_tree.heading(col, text=col.replace('_', ' ').title())
        
        # Run load function
        self.load_capex_summary()

    def load_capex_summary(self):
        """Load and display expenditure summary"""
        for item in self.overview_tree.get_children():
            self.overview_tree.delete(item)

        # Configure the tags (Here I import the colour settings for each tag name, assign a variable and set the tree to take the colour)
        capex_highlights = ASSET_SETTINGS.get('capex_highlights', [])
        for ch in capex_highlights:
            tag_name = ch.get('status', '')
            colour = ch.get('colour', '#FFFFFF')
            if tag_name:
                self.overview_tree.tag_configure(tag_name, background=colour)

        try:
            with open(ASSET_SETTINGS_PATH, 'r') as s:
                asset_settings = json.load(s)

            for gl in asset_settings.get('capexs', []):
                code = gl.get('code', '')
                name = gl.get('name', '')
                total_budget = float(gl.get('max_amount', 0.0))

                # Get currency symbol from settings - use hardcoded symbols to avoid encoding issues
                selected_currency = asset_settings.get('currency', 'GBP')
                # Map currency codes to symbols directly to avoid encoding issues
                currency_map = {
                    'GBP': '\u00A3',  # £
                    'EUR': '\u20AC',  # €
                    'USD': '$'
                }
                currency_symbol = currency_map.get(selected_currency, '\u00A3')
                used_budget = 0.0
                num_costs = 0
                
                # Calculate used budget and number of costs from costs.json
                if COST_PATH.exists():
                    with open(COST_PATH, 'r') as c:
                        costs_data = json.load(c)
                        for cost in costs_data.get('costs', []):
                            lines = cost.get('line_items', [])
                            cost_has_matching_gl = False
                            for line in lines:
                                if line.get('gl_code') == code:
                                    amount = float(line.get('total_cost', 0.0))
                                    used_budget += amount
                                    num_costs += 1
                                    cost_has_matching_gl = True
                            # Only add carriage if this cost has at least one line with matching GL code
                            if cost_has_matching_gl:
                                carriage = float(cost.get('carriage', 0.0))
                                used_budget += carriage
                
                remaining_budget = total_budget - used_budget
                percentage_used = (used_budget / total_budget * 100) if total_budget > 0 else 0.0
                
                # Determine tag based on rules - sort by threshold descending to apply highest matching threshold
                tag = ()
                sorted_highlights = sorted(capex_highlights, key=lambda x: float(x.get('threshold', 0.0)), reverse=True)
                for ch in sorted_highlights:
                    threshold = float(ch.get('threshold', 0.0))
                    
                    if percentage_used >= threshold:
                        tag = (ch.get('status', ''),)
                        break  # Apply the first matching tag only

                self.overview_tree.insert('', 'end', values=(
                    code,
                    name,
                    f"{currency_symbol}{total_budget:,.2f}",
                    num_costs,
                    f"{currency_symbol}{remaining_budget:,.2f}",
                    f"{currency_symbol}{used_budget:,.2f}",
                    f"{percentage_used:.2f}%"
                ), tags=tag)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load expenditure summary: {str(e)}")

    def load_assets(self):
        """Load merged assets from both index (read-only) and registry (editable), sorted by next_cal_due"""
        for item in self.assets_tree.get_children():
            self.assets_tree.delete(item)
        
        try:
            # Load both data sources directly
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            # Create a merged view
            # Build dictionaries keyed by identifier for easy merging
            index_dict = {item.get('identifier', ''): item for item in index_data} if index_data else {}
            registry_dict = {item.get('identifier', ''): item for item in registry_data} if registry_data else {}
            
            # Get all unique identifiers
            all_identifiers = set(list(index_dict.keys()) + list(registry_dict.keys()))
            all_identifiers.discard('')  # Remove empty identifier if present
            
            # Determine all columns from both sources
            index_cols = []
            registry_cols = ['identifier', 'cal_date', 'cal_period', 'cal_by', 'next_cal_due', 
                           'check_period', 'last_checked', 'location', 'status']
            
            if index_data and len(index_data) > 0:
                index_cols = [col for col in index_data[0].keys() if col != 'identifier']
            
            # Combine columns: identifier first, then index columns, then registry columns
            all_columns = ['identifier'] + index_cols + [col for col in registry_cols if col != 'identifier']
            
            # Filter out ignored columns from settings
            ignore_cols = ASSET_SETTINGS.get('ignore_columns', [])
            displayed_columns = [col for col in all_columns if col not in ignore_cols]
            
            # Configure treeview columns (only displayed ones)
            self.assets_tree['columns'] = displayed_columns
            self.assets_tree.column('#0', width=0, stretch=False)
            
            for col in displayed_columns:
                self.assets_tree.column(col, width=120, minwidth=80)
                # Mark registry columns as editable in the header
                if col in registry_cols and col != 'identifier':
                    self.assets_tree.heading(col, text=col.replace('_', ' ').title() + ' (E)')
                else:
                    self.assets_tree.heading(col, text=col.replace('_', ' ').title())
            
            # Create search entries if they don't exist
            if not self.assets_search_vars:
                for idx, col in enumerate(displayed_columns):
                    self.assets_search_frame.grid_columnconfigure(idx, weight=1)
                    
                    ttk.Label(self.assets_search_frame, text=col.replace('_', ' ').title(), 
                             font=('TkDefaultFont', 8)).grid(row=0, column=idx, sticky='ew', padx=2)
                    
                    search_var = tk.StringVar()
                    self.assets_search_vars[col] = search_var
                    
                    entry = ttk.Entry(self.assets_search_frame, textvariable=search_var)
                    entry.grid(row=1, column=idx, sticky='ew', padx=2, pady=2)
                    entry.bind('<KeyRelease>', lambda e: self.filter_assets())
            
            # Helper function to format dates to DD/MM/YYYY
            def format_date(date_str):
                if not date_str:
                    return ''
                # Try different date formats
                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d']:
                    try:
                        dt = datetime.strptime(str(date_str).strip(), fmt)
                        return dt.strftime('%d/%m/%Y')
                    except:
                        continue
                # If no format matches, return as-is
                return str(date_str)
            
            # Collect all rows for sorting
            rows = []
            for identifier in all_identifiers:
                index_item = index_dict.get(identifier, {})
                registry_item = registry_dict.get(identifier, {})
                
                # Build row values
                row_data = {}
                for col in all_columns:
                    if col in registry_cols:
                        value = registry_item.get(col, '')
                    else:
                        value = index_item.get(col, '')
                    
                    # Format date columns
                    if 'date' in col.lower() or col in ['cal_date', 'next_cal_due', 'last_checked']:
                        value = format_date(value)
                    
                    row_data[col] = value
                
                rows.append(row_data)
            
            # Sort by next_cal_due with Classification and Cal_period filters
            def cal_due_sort_key(row):
                classification = str(row.get('Classification', '')).strip()
                cal_period = str(row.get('cal_period', '')).strip()
                cal_due = row.get('next_cal_due', '')
                
                # Send to bottom if Classification is Obsolete or cal_period is None/empty
                if classification.lower() == 'obsolete' or not cal_period or cal_period.lower() == 'none':
                    return (2, '')
                
                # Empty cal_due dates go to middle (between normal and obsolete)
                if not cal_due:
                    return (1, '')
                
                # Normal items sorted by date
                try:
                    return (0, datetime.strptime(cal_due, '%d/%m/%Y'))
                except:
                    return (1, cal_due)
            
            rows.sort(key=cal_due_sort_key)
            
            # Display sorted rows (only displayed columns)
            for row_data in rows:
                values = [row_data.get(col, '') for col in displayed_columns]
                self.assets_tree.insert('', 'end', values=values)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load assets: {str(e)}")
    
    def filter_assets(self):
        """Filter merged assets view based on column-specific search, sorted by cal due date (oldest first)"""
        for item in self.assets_tree.get_children():
            self.assets_tree.delete(item)
        
        try:
            # Load both data sources directly
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            # Create dictionaries
            index_dict = {item.get('identifier', ''): item for item in index_data} if index_data else {}
            registry_dict = {item.get('identifier', ''): item for item in registry_data} if registry_data else {}
            
            # Get all unique identifiers
            all_identifiers = set(list(index_dict.keys()) + list(registry_dict.keys()))
            all_identifiers.discard('')
            
            # Get columns and search filters
            displayed_columns = self.assets_tree['columns']
            registry_cols = ['identifier', 'cal_date', 'cal_period', 'cal_by', 'next_cal_due', 
                           'check_period', 'last_checked', 'location', 'status']
            
            # Get all possible columns (including ignored ones) for complete data loading
            index_cols = []
            if index_data and len(index_data) > 0:
                index_cols = [col for col in index_data[0].keys() if col != 'identifier']
            all_columns = ['identifier'] + index_cols + [col for col in registry_cols if col != 'identifier']
            
            filters = {col: self.assets_search_vars.get(col, tk.StringVar()).get().lower() for col in displayed_columns}
            
            # Helper function to format dates to DD/MM/YYYY
            def format_date(date_str):
                if not date_str:
                    return ''
                # Try different date formats
                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d']:
                    try:
                        dt = datetime.strptime(str(date_str).strip(), fmt)
                        return dt.strftime('%d/%m/%Y')
                    except:
                        continue
                # If no format matches, return as-is
                return str(date_str)
            
            # Collect matching rows
            matching_rows = []
            
            for identifier in all_identifiers:
                index_item = index_dict.get(identifier, {})
                registry_item = registry_dict.get(identifier, {})
                
                # Build combined data for this row
                row_data = {}
                for col in all_columns:
                    if col in registry_cols:
                        value = registry_item.get(col, '')
                    else:
                        value = index_item.get(col, '')
                    
                    # Format date columns
                    if 'date' in col.lower() or col in ['cal_date', 'next_cal_due', 'last_checked']:
                        value = format_date(value)
                    
                    row_data[col] = value
                
                # Check if row matches all filters
                match = True
                for col, search_text in filters.items():
                    if search_text:
                        value = str(row_data.get(col, '')).lower()
                        if search_text not in value:
                            match = False
                            break
                
                if match:
                    matching_rows.append(row_data)
            
            # Sort by next_cal_due with Classification and Cal_period filters
            def cal_due_sort_key(row):
                classification = str(row.get('Classification', '')).strip()
                cal_period = str(row.get('cal_period', '')).strip()
                cal_due = row.get('next_cal_due', '')
                
                # Send to bottom if Classification is Obsolete or cal_period is None/empty
                if classification.lower() == 'obsolete' or not cal_period or cal_period.lower() == 'none':
                    return (2, '')
                
                # Empty cal_due dates go to middle (between normal and obsolete)
                if not cal_due:
                    return (1, '')
                
                # Normal items sorted by date
                try:
                    return (0, datetime.strptime(cal_due, '%d/%m/%Y'))
                except:
                    return (1, cal_due)
            
            matching_rows.sort(key=cal_due_sort_key)
            
            # Display sorted rows (only displayed columns)
            for row_data in matching_rows:
                values = [row_data.get(col, '') for col in displayed_columns]
                self.assets_tree.insert('', 'end', values=values)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to filter assets: {str(e)}")

    def edit_selected_ticket(self):
        selection = self.ticket_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a ticket to edit")
            return
        
        item = self.ticket_tree.item(selection[0])
        ticket_id = item['values'][0]
        
        # Load ticket details
        try:
            with open(TIC_PATH, 'r') as f:
                data = json.load(f)
            
            ticket = next((t for t in data.get('tickets', []) if t.get('ticket_id') == ticket_id), None)
            if ticket:
                EditTicketWindow(self.root, self.current_user.get(), self.load_tickets, ticket_data=ticket)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load ticket for editing: {str(e)}")

    def load_tickets(self):
        """Load and display tickets"""
        for item in self.ticket_tree.get_children():
            self.ticket_tree.delete(item)
        
        try:
            with open(TIC_PATH, 'r') as f:
                data = json.load(f)
            
            # Load costs data
            cost_data = {}
            if COST_PATH.exists():
                with open(COST_PATH, 'r') as f:
                    costs_json = json.load(f)
                    # Build a dict of ticket_id -> list of PO numbers
                    for cost in costs_json.get('costs', []):
                        ticket_id = cost.get('ticket_id')
                        po_num = cost.get('po_number', '')
                        if ticket_id:
                            if ticket_id not in cost_data:
                                cost_data[ticket_id] = []
                            if po_num:
                                cost_data[ticket_id].append(po_num)
            
            tickets = data.get('tickets', [])
            for ticket in tickets:
                # Get first asset or "Multiple" if more than one
                assets = ticket.get('assets', [])
                asset_display = assets[0] if len(assets) == 1 else f"Multiple ({len(assets)})" if len(assets) > 1 else "None"
                
                # Format date to DD/MM/YYYY
                created_date = ticket.get('created_date', '')
                if created_date:
                    try:
                        dt = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                        created_date = dt.strftime('%d/%m/%Y')
                    except:
                        created_date = created_date.split('T')[0] if 'T' in created_date else created_date
                
                # Get PO numbers for this ticket
                ticket_id = ticket.get('ticket_id', '')
                po_numbers = cost_data.get(ticket_id, [])
                po_display = ', '.join(po_numbers) if po_numbers else ''

                self.ticket_tree.insert('', 'end', values=(
                    ticket_id,
                    ticket.get('title', ''),
                    ticket.get('type', ''),
                    asset_display,
                    created_date,
                    ticket.get('created_by', ''),
                    ticket.get('status', 'Open'),
                    po_display
                ))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load tickets: {str(e)}")
    
    def filter_tickets(self):
        """Filter tickets based on column-specific search"""
        for item in self.ticket_tree.get_children():
            self.ticket_tree.delete(item)
        
        try:
            with open(TIC_PATH, 'r') as f:
                data = json.load(f)
            
            # Load costs data for PO number filtering
            cost_data = {}
            if COST_PATH.exists():
                with open(COST_PATH, 'r') as f:
                    costs_json = json.load(f)
                    # Build a dict of ticket_id -> list of PO numbers
                    for cost in costs_json.get('costs', []):
                        ticket_id = cost.get('ticket_id')
                        po_num = cost.get('po_number', '')
                        if ticket_id:
                            if ticket_id not in cost_data:
                                cost_data[ticket_id] = []
                            if po_num:
                                cost_data[ticket_id].append(po_num)
            
            # Get search filters
            filters = {col: self.ticket_search_vars.get(col, tk.StringVar()).get().lower() 
                      for col in ['ticket_id', 'type', 'asset', 'date', 'user', 'status', 'po_number']}
            
            tickets = data.get('tickets', [])
            for ticket in tickets:
                # Display all assets, comma-separated
                assets = ticket.get('assets', [])
                asset_display = ', '.join(assets) if assets else "None"
                
                # Format date to DD/MM/YYYY
                created_date = ticket.get('created_date', '')
                if created_date:
                    try:
                        dt = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                        created_date = dt.strftime('%d/%m/%Y')
                    except:
                        created_date = created_date.split('T')[0] if 'T' in created_date else created_date
                
                # Get PO numbers for this ticket
                ticket_id = ticket.get('ticket_id', '')
                po_numbers = cost_data.get(ticket_id, [])
                po_display = ', '.join(po_numbers) if po_numbers else ''
                
                # Build row data
                row_data = {
                    'ticket_id': ticket_id,
                    'type': ticket.get('type', ''),
                    'asset': asset_display,
                    'date': created_date,
                    'user': ticket.get('created_by', ''),
                    'status': ticket.get('status', 'Open'),
                    'po_number': po_display
                }
                
                # Check if ticket matches all filters
                match = True
                for col, search_text in filters.items():
                    if search_text:
                        value = str(row_data.get(col, '')).lower()
                        if search_text not in value:
                            match = False
                            break
                
                if match:
                    self.ticket_tree.insert('', 'end', values=(
                        row_data['ticket_id'],
                        row_data['type'],
                        row_data['asset'],
                        row_data['date'],
                        row_data['user'],
                        row_data['status'],
                        row_data['po_number']
                    ))
        except Exception as e:
            pass  # Silently ignore filter errors
    
    def create_new_ticket(self):
        try:
            NewTicketWindow(self.root, self.current_user.get(), self.load_tickets, preset_assets=[])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create ticket window: {str(e)}")
    
    def on_ticket_select(self, event):
        selection = self.ticket_tree.selection()
        if not selection:
            return
        
        item = self.ticket_tree.item(selection[0])
        ticket_id = item['values'][0]
        
        # Clear detail frame
        for widget in self.detail_frame.winfo_children():
            widget.destroy()
        
        # Load ticket details
        try:
            with open(TIC_PATH, 'r') as f:
                data = json.load(f)
            
            ticket = next((t for t in data.get('tickets', []) if t.get('ticket_id') == ticket_id), None)
            if ticket:
                self.display_ticket_detail(ticket)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load ticket details: {str(e)}")
    
    def display_ticket_detail(self, ticket):
        # Header
        header_frame = ttk.Frame(self.detail_frame)
        header_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(header_frame, text=f"Ticket: {ticket.get('ticket_id', 'N/A')}", 
                 font=('TkDefaultFont', 14, 'bold')).pack(anchor='w')
        
        # Display assets
        assets_display = ', '.join(ticket.get('assets', [])) if isinstance(ticket.get('assets'), list) else ticket.get('asset', 'N/A')
        ttk.Label(header_frame, text=f"Type: {ticket.get('type', 'N/A')} | Assets: {assets_display}").pack(anchor='w')
        ttk.Label(header_frame, text=f"User: {ticket.get('created_by', 'N/A')} | Date: {ticket.get('created_date', 'N/A')} | Status: {ticket.get('status', 'Open')}").pack(anchor='w')
        ttk.Label(header_frame, text=f"Title: {ticket.get('title', 'N/A')}").pack(anchor='w', pady=(5,0))
        ttk.Label(header_frame, text=f"Description: {ticket.get('description', 'No Description Entered')}").pack(anchor='w', pady=(0,5))
        
        # Show attached files if any
        if ticket.get('attached_files'):
            files_frame = ttk.LabelFrame(header_frame, text="Attached Files", padding=5)
            files_frame.pack(fill='x', pady=5)
            
            files_list = tk.Listbox(files_frame, height=3)
            files_list.pack(side='left', fill='both', expand=True)
            
            for file_info in ticket.get('attached_files', []):
                files_list.insert('end', file_info.get('filename', 'Unknown'))
            
            files_btn_frame = ttk.Frame(files_frame)
            files_btn_frame.pack(side='right', fill='y', padx=5)
            
            def open_ticket_file():
                selection = files_list.curselection()
                if selection:
                    file_info = ticket.get('attached_files', [])[selection[0]]
                    file_path = file_info.get('path', '')
                    if file_path and Path(file_path).exists():
                        try:
                            os.startfile(file_path)
                        except Exception as e:
                            messagebox.showerror("Error", f"Failed to open file: {str(e)}")
                    else:
                        messagebox.showerror("Error", "File not found")
            
            def remove_ticket_file():
                selection = files_list.curselection()
                if not selection:
                    messagebox.showwarning("Warning", "Please select a file to remove")
                    return
                
                if messagebox.askyesno("Confirm", "Remove this file from the ticket?"):
                    idx = selection[0]
                    try:
                        with open(TIC_PATH, 'r') as f:
                            ticket_data = json.load(f)
                        
                        for t in ticket_data.get('tickets', []):
                            if t.get('ticket_id') == ticket.get('ticket_id'):
                                attached_files = t.get('attached_files', [])
                                if idx < len(attached_files):
                                    attached_files.pop(idx)
                                    t['attached_files'] = attached_files
                                break
                        
                        save_json(TIC_PATH, ticket_data)
                        
                        self.on_ticket_select(None)
                        messagebox.showinfo("Success", "File removed")
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to remove file: {str(e)}")
            
            ttk.Button(files_btn_frame, text="Open", command=open_ticket_file).pack(fill='x', pady=2)
            ttk.Button(files_btn_frame, text="Remove", command=remove_ticket_file).pack(fill='x', pady=2)
        
        # Show closed info if ticket is closed
        if ticket.get('status') == 'Closed':
            closed_text = f"Closed: {ticket.get('closed_date', 'N/A')} by {ticket.get('closed_by', 'N/A')}"
            ttk.Label(header_frame, text=closed_text, font=('TkDefaultFont', 9, 'italic'), foreground='red').pack(anchor='w')
        
        # Buttons
        btn_frame = ttk.Frame(self.detail_frame)
        btn_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(btn_frame, text="Add Action", command=lambda: self.add_action(ticket)).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Add Cost", command=lambda: self.add_cost(ticket)).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Update Assets", command=lambda: self.update_assets(ticket)).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Attach File", command=lambda: self.import_file_to_ticket(ticket)).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Close Ticket", command=lambda: self.close_ticket(ticket)).pack(side='left', padx=2)
        
        # Actions
        ttk.Label(self.detail_frame, text="Actions:", font=('TkDefaultFont', 11, 'bold')).pack(anchor='w', padx=10, pady=(10, 5))
        
        actions_frame = ttk.Frame(self.detail_frame)
        actions_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        actions_scroll = ttk.Scrollbar(actions_frame, orient='vertical')
        self.actions_tree = ttk.Treeview(actions_frame, yscrollcommand=actions_scroll.set, height=10)
        actions_scroll.config(command=self.actions_tree.yview)
        
        actions_scroll.pack(side='right', fill='y')
        self.actions_tree.pack(fill='both', expand=True)
        
        self.actions_tree['columns'] = ('action_id', 'type', 'user', 'date', 'ref_assets', 'description')
        self.actions_tree.column('#0', width=0, stretch=False)
        
        for col in self.actions_tree['columns']:
            self.actions_tree.column(col, width=100)
            self.actions_tree.heading(col, text=col.replace('_', ' ').title())
        
        for action in ticket.get('actions', []):
            ref_assets = action.get('referenced_assets', [])
            ref_assets_str = ', '.join(ref_assets) if ref_assets else 'All'
            
            self.actions_tree.insert('', 'end', values=(
                action.get('action_id', ''),
                action.get('type', ''),
                action.get('user', ''),
                action.get('date', ''),
                ref_assets_str,
                action.get('description', '')[:50]
            ))
        
        # Bind double-click to edit action
        self.actions_tree.bind('<Double-1>', lambda e: self.edit_action(ticket))
        
        # Costs
        ttk.Label(self.detail_frame, text="Costs:", font=('TkDefaultFont', 11, 'bold')).pack(anchor='w', padx=10, pady=(10, 5))
        
        costs_frame = ttk.Frame(self.detail_frame)
        costs_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        costs_scroll = ttk.Scrollbar(costs_frame, orient='vertical')
        self.costs_tree = ttk.Treeview(costs_frame, yscrollcommand=costs_scroll.set, height=8)
        costs_scroll.config(command=self.costs_tree.yview)
        
        costs_scroll.pack(side='right', fill='y')
        self.costs_tree.pack(fill='both', expand=True)
        
        self.costs_tree['columns'] = ('cost_id', 'po_number', 'supplier', 'total_cost', 'currency', 'date', 'signoff')
        self.costs_tree.column('#0', width=0, stretch=False)
        
        for col in self.costs_tree['columns']:
            self.costs_tree.column(col, width=100)
            self.costs_tree.heading(col, text=col.replace('_', ' ').title())
        
        # Load costs from costs.json that are linked to this ticket
        try:
            if COST_PATH.exists():
                with open(COST_PATH, 'r') as f:
                    cost_data = json.load(f)
                
                for cost in cost_data.get('costs', []):
                    if cost.get('ticket_id') == ticket.get('ticket_id'):
                        self.costs_tree.insert('', 'end', values=(
                            cost.get('cost_id', ''),
                            cost.get('po_number', ''),
                            cost.get('supplier', ''),
                            cost.get('final_cost', ''),
                            cost.get('currency', ''),
                            cost.get('date', ''),
                            cost.get('signoff', '')
                        ))
        except Exception as e:
            print(f"Error loading costs: {e}")
        
        self.costs_tree.bind('<Double-1>', lambda e: self.view_cost_detail(ticket))
    
    def add_action(self, ticket):
        AddActionWindow(self.root, ticket, self.current_user.get(), lambda: self.on_ticket_select(None))
    
    def edit_action(self, ticket):
        """Edit the selected action"""
        selection = self.actions_tree.selection()
        if not selection:
            return
        
        item = self.actions_tree.item(selection[0])
        action_id = item['values'][0]
        
        # Find the action
        action = next((a for a in ticket.get('actions', []) if a.get('action_id') == action_id), None)
        if action:
            AddActionWindow(self.root, ticket, self.current_user.get(), lambda: self.on_ticket_select(None), action_to_edit=action)
    
    def add_cost(self, ticket):
        AddCostWindow(self.root, ticket, self.current_user.get(), lambda: self.on_ticket_select(None))
    
    def update_assets(self, ticket):
        UpdateAssetsWindow(self.root, ticket, self.current_user.get(), lambda: self.on_ticket_select(None))
    
    def view_cost_detail(self, ticket):
        """Open cost for editing when double-clicked"""
        selection = self.costs_tree.selection()
        if not selection:
            return
        
        item = self.costs_tree.item(selection[0])
        values = item['values']
        
        if not values or len(values) < 1:
            return
        
        cost_id = values[0]
        
        # Load cost from costs.json
        try:
            if not COST_PATH.exists():
                messagebox.showwarning("Not Found", "Costs file not found")
                return
            
            with open(COST_PATH, 'r') as f:
                cost_data = json.load(f)
            
            # Find the cost by cost_id
            cost = None
            for c in cost_data.get('costs', []):
                if c.get('cost_id') == cost_id:
                    cost = c
                    break
            
            if cost:
                # Open in edit mode
                AddCostWindow(self.root, ticket, self.current_user.get(), lambda: self.on_ticket_select(None), cost_to_edit=cost)
            else:
                messagebox.showwarning("Not Found", f"Could not find cost with ID: {cost_id}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load cost: {str(e)}")
    
    def close_ticket(self, ticket):
        # Check if all dispatched assets have been returned
        try:
            # Load registry to check asset statuses
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            # Get assets from ticket
            assets = ticket.get('assets', [])
            if not isinstance(assets, list):
                assets = [ticket.get('asset', '')]
            
            # Check if any assets are still External
            external_assets = []
            for asset_id in assets:
                reg_item = next((r for r in registry_data if r.get('identifier') == asset_id), None)
                if reg_item and reg_item.get('status') == 'External':
                    external_assets.append(asset_id)
            
            if external_assets:
                messagebox.showwarning("Cannot Close Ticket", 
                    f"The following assets are still External and must be returned before closing:\n" + 
                    ', '.join(external_assets))
                return
            
            if messagebox.askyesno("Close Ticket", f"Close ticket {ticket.get('ticket_id')}?"):
                # Load ticket data
                with open(TIC_PATH, 'r') as f:
                    ticket_data = json.load(f)
                
                for t in ticket_data.get('tickets', []):
                    if t.get('ticket_id') == ticket.get('ticket_id'):
                        t['status'] = 'Closed'
                        t['closed_date'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                        t['closed_by'] = self.current_user.get()
                        break
                
                # Save to temp then main
                save_json(TIC_PATH, ticket_data)
                
                self.load_tickets()
                messagebox.showinfo("Success", "Ticket closed successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to close ticket: {str(e)}")
    
    def import_file_to_ticket(self, ticket):
        file_path = filedialog.askopenfilename(title="Select file to attach")
        if file_path:
            try:
                # Copy file to ticket folder
                copied_path = copy_file_to_ticket(file_path, ticket.get('ticket_id'))
                
                if not copied_path:
                    messagebox.showerror("Error", "Failed to copy file to ticket folder")
                    return
                
                # Load ticket data
                with open(TIC_PATH, 'r') as f:
                    ticket_data = json.load(f)
                
                for t in ticket_data.get('tickets', []):
                    if t.get('ticket_id') == ticket.get('ticket_id'):
                        # Add to attached files with new path
                        attached_files = t.get('attached_files', [])
                        attached_files.append({
                            'filename': Path(copied_path).name,
                            'path': copied_path,
                            'uploaded_by': self.current_user.get(),
                            'uploaded_date': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                        })
                        t['attached_files'] = attached_files
                        
                        # Add action for file attachment
                        actions = t.get('actions', [])
                        action_num = len(actions) + 1
                        action_id = f"{ticket.get('ticket_id')}-{action_num}"
                        
                        actions.append({
                            'action_id': action_id,
                            'type': 'File Attachment',
                            'user': self.current_user.get(),
                            'date': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                            'description': f"Attached file: {Path(copied_path).name}"
                        })
                        t['actions'] = actions
                        break
                
                # Save with backup
                save_json(TIC_PATH, ticket_data)
                
                self.on_ticket_select(None)
                messagebox.showinfo("Success", "File attached successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to attach file: {str(e)}")
    
    def new_ticket_from_asset(self):
        selection = self.assets_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select at least one asset")
            return
        
        # Get selected asset IDs
        asset_ids = [self.assets_tree.item(item)['values'][0] for item in selection]
        NewTicketWindow(self.root, self.current_user.get(), self.load_tickets, preset_assets=asset_ids)
    
    def view_asset_tickets(self):
        selection = self.assets_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an asset")
            return
        
        asset_id = self.assets_tree.item(selection[0])['values'][0]
        ViewAssetTicketsWindow(self.root, asset_id)
    
    def change_asset_location(self):
        selection = self.assets_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select at least one asset")
            return
        
        asset_ids = [self.assets_tree.item(item)['values'][0] for item in selection]
        ChangeLocationWindow(self.root, asset_ids, self.current_user.get(), self.load_assets)
    
    def change_asset_status(self):
        selection = self.assets_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select at least one asset")
            return
        
        asset_ids = [self.assets_tree.item(item)['values'][0] for item in selection]
        ChangeStatusWindow(self.root, asset_ids, self.current_user.get(), self.load_assets)
    
    def view_asset_details(self):
        selection = self.assets_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an asset")
            return
        
        asset_id = self.assets_tree.item(selection[0])['values'][0]
        AssetDetailsWindow(self.root, asset_id)
    
    def edit_registry_item(self):
        selection = self.assets_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an asset to edit registry data")
            return
        
        item = self.assets_tree.item(selection[0])
        identifier = item['values'][0]
        
        try:
            # Load registry
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            # Find or create registry entry for this identifier
            registry_item = next((r for r in registry_data if r.get('identifier') == identifier), None)
            
            if not registry_item:
                # Create new registry entry if doesn't exist
                registry_item = {
                    'identifier': identifier,
                    'cal_date': '',
                    'cal_period': '',
                    'cal_by': '',
                    'next_cal_due': '',
                    'check_period': '',
                    'last_checked': '',
                    'location': ''
                }
                registry_data.append(registry_item)
                # Save with backup
                save_json(REG_PATH, registry_data)
            
            EditRegistryWindow(self.root, registry_item, self.load_assets)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load item: {str(e)}")
    
    def edit_asset_settings(self):
        """Open a live-editing settings window"""
        SettingsEditorWindow(self.root, ASSET_SETTINGS_PATH)
    
    def add_new_asset(self):
        """Add a new asset to both index and registry"""
        AddAssetWindow(self.root, self.load_assets)
    
    def reload_data(self):
        self.load_assets()
        self.load_tickets()
        messagebox.showinfo("Success", "Data reloaded")

class UpdateAssetsWindow:
    def __init__(self, parent, ticket, user, callback):
        self.ticket = ticket
        self.user = user
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title(f"Update Assets - {ticket.get('ticket_id')}")
        self.window.geometry("800x600")
        self.window.attributes('-topmost', True)
        
        # Get assets from ticket
        assets = ticket.get('assets', [])
        if not isinstance(assets, list):
            # Handle old format where asset was a string
            assets = [ticket.get('asset', '')]
        
        if not assets or not any(assets):
            messagebox.showwarning("Warning", "No assets found in this ticket")
            self.window.destroy()
            return
        
        # Create notebook with tab for each asset
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.asset_editors = {}
        
        # Load registry data - check temp first
        with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
        
        self.registry_data = registry_data
        
        # Create tab for each asset
        for asset_id in assets:
            if asset_id:
                self.create_asset_tab(asset_id)
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(fill='x', padx=10, pady=10)
        ttk.Button(btn_frame, text="Save All Changes", command=self.save_changes).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def create_asset_tab(self, asset_id):
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=asset_id)
        
        # Find registry item for this asset
        registry_item = next((r for r in self.registry_data if r.get('identifier') == asset_id), None)
        
        if not registry_item:
            # Create new registry entry
            registry_item = {
                'identifier': asset_id,
                'cal_date': '',
                'cal_period': '',
                'cal_by': '',
                'next_cal_due': '',
                'check_period': '',
                'last_checked': '',
                'location': '',
                'status': 'Available'
            }
            self.registry_data.append(registry_item)
        
        # Create form fields
        entries = {}
        row = 0
        
        # Asset ID (read-only)
        ttk.Label(tab_frame, text="Asset ID:", font=('TkDefaultFont', 10, 'bold')).grid(row=row, column=0, sticky='w', padx=10, pady=10)
        ttk.Label(tab_frame, text=asset_id).grid(row=row, column=1, sticky='w', padx=10, pady=10)
        row += 1
        
        # Editable fields
        registry_fields = ['cal_date', 'cal_period', 'cal_by', 'next_cal_due', 
                          'check_period', 'last_checked', 'location', 'status']
        
        for field in registry_fields:
            ttk.Label(tab_frame, text=f"{field.replace('_', ' ').title()}:", 
                     font=('TkDefaultFont', 10)).grid(row=row, column=0, sticky='w', padx=10, pady=5)
            
            if field == 'location':
                entry = ttk.Combobox(tab_frame, values=ASSET_SETTINGS.get('locations', []), width=40)
                entry.set(registry_item.get(field, ''))
            elif field == 'cal_by':
                entry = ttk.Entry(tab_frame, width=42)
                entry.insert(0, registry_item.get(field, ''))
            elif field == 'status':
                entry = ttk.Combobox(tab_frame, values=ASSET_SETTINGS.get('asset_statuses', []), width=40)
                entry.set(registry_item.get(field, 'Available'))
            else:
                entry = ttk.Entry(tab_frame, width=42)
                entry.insert(0, registry_item.get(field, ''))
            
            entry.grid(row=row, column=1, padx=10, pady=5, sticky='w')
            entries[field] = entry
            row += 1
        
        self.asset_editors[asset_id] = {'registry_item': registry_item, 'entries': entries}
    
    def save_changes(self):
        try:
            # Update registry items with new values
            for asset_id, editor in self.asset_editors.items():
                registry_item = editor['registry_item']
                entries = editor['entries']
                
                for field, entry in entries.items():
                    registry_item[field] = entry.get()
            
            # Save with backup
            save_json(REG_PATH, self.registry_data)
            
            # Add action to ticket
            with open(TIC_PATH, 'r') as f:
                ticket_data = json.load(f)
            
            for t in ticket_data.get('tickets', []):
                if t.get('ticket_id') == self.ticket.get('ticket_id'):
                    actions = t.get('actions', [])
                    action_num = len(actions) + 1
                    action_id = f"{self.ticket.get('ticket_id')}-{action_num}"
                    
                    assets_updated = ', '.join(self.asset_editors.keys())
                    actions.append({
                        'action_id': action_id,
                        'type': 'Asset Update',
                        'user': self.user,
                        'date': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                        'description': f"Updated registry data for assets: {assets_updated}"
                    })
                    t['actions'] = actions
                    break
            
            # Save ticket data to temp then main
            save_json(TIC_PATH, ticket_data)
            
            self.callback()
            self.window.destroy()
            messagebox.showinfo("Success", "Asset registry data updated successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save changes: {str(e)}")

class NewTicketWindow:
    def __init__(self, parent, user, callback, preset_assets=[]):
        self.user = user
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("Create New Ticket")
        self.window.geometry("800x700")
        self.window.attributes('-topmost', True)
        
        # Title
        ttk.Label(self.window, text="Create New Ticket", 
                 font=('TkDefaultFont', 14, 'bold')).pack(pady=10)
        
        # Main frame
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # Ticket Type
        ttk.Label(main_frame, text="Ticket Type:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=0, column=0, sticky='w', pady=5)
        self.ticket_type = ttk.Combobox(main_frame, values=TICKET_TYPES, 
                                        state='readonly', width=30)
        self.ticket_type.grid(row=0, column=1, sticky='ew', pady=5)
        self.ticket_type.set("General")
        
        # Title
        ttk.Label(main_frame, text="Title:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=1, column=0, sticky='w', pady=5)
        self.title_entry = ttk.Entry(main_frame, width=50)
        self.title_entry.grid(row=1, column=1, sticky='ew', pady=5)
        
        # Description
        ttk.Label(main_frame, text="Description:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=2, column=0, sticky='nw', pady=5)
        self.description_text = tk.Text(main_frame, height=4, width=50)
        self.description_text.grid(row=2, column=1, sticky='ew', pady=5)
        
        # Assets
        ttk.Label(main_frame, text="Assets:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=3, column=0, sticky='nw', pady=5)
        
        asset_frame = ttk.Frame(main_frame)
        asset_frame.grid(row=3, column=1, sticky='ew', pady=5)
        
        # Asset listbox
        self.asset_listbox = tk.Listbox(asset_frame, height=5)
        self.asset_listbox.pack(side='left', fill='both', expand=True)
        
        asset_scroll = ttk.Scrollbar(asset_frame, orient='vertical', 
                                     command=self.asset_listbox.yview)
        asset_scroll.pack(side='right', fill='y')
        self.asset_listbox.config(yscrollcommand=asset_scroll.set)
        
        # Add preset assets
        for asset in preset_assets:
            self.asset_listbox.insert('end', asset)
        
        # Asset buttons
        asset_btn_frame = ttk.Frame(main_frame)
        asset_btn_frame.grid(row=4, column=1, sticky='w', pady=5)
        ttk.Button(asset_btn_frame, text="Add Asset", command=self.add_asset).pack(side='left', padx=2)
        ttk.Button(asset_btn_frame, text="Remove Selected", command=self.remove_asset).pack(side='left', padx=2)
        
        main_frame.grid_columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Create Ticket", command=self.create_ticket).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def add_asset(self):
        """Add an asset to the ticket"""
        # Get all available assets
        try:
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            
            asset_ids = [item.get('identifier', '') for item in index_data if item.get('identifier')]
            
            if not asset_ids:
                messagebox.showwarning("No Assets", "No assets found in index.")
                return
            
            # Simple selection dialog
            select_window = tk.Toplevel(self.window)
            select_window.title("Select Asset")
            select_window.geometry("400x500")
            
            ttk.Label(select_window, text="Select Asset:", font=('TkDefaultFont', 10, 'bold')).pack(pady=10)
            
            # Search
            search_var = tk.StringVar()
            search_entry = ttk.Entry(select_window, textvariable=search_var)
            search_entry.pack(fill='x', padx=10, pady=5)
            
            # Listbox
            listbox_frame = ttk.Frame(select_window)
            listbox_frame.pack(fill='both', expand=True, padx=10, pady=5)
            
            asset_list = tk.Listbox(listbox_frame)
            asset_list.pack(side='left', fill='both', expand=True)
            
            scroll = ttk.Scrollbar(listbox_frame, orient='vertical', command=asset_list.yview)
            scroll.pack(side='right', fill='y')
            asset_list.config(yscrollcommand=scroll.set)
            
            def filter_assets(*args):
                asset_list.delete(0, 'end')
                search_term = search_var.get().lower()
                for asset_id in sorted(asset_ids):
                    if search_term in asset_id.lower():
                        asset_list.insert('end', asset_id)
            
            search_var.trace('w', filter_assets)
            filter_assets()
            
            def select_asset():
                selection = asset_list.curselection()
                if selection:
                    selected = asset_list.get(selection[0])
                    if selected not in self.asset_listbox.get(0, 'end'):
                        self.asset_listbox.insert('end', selected)
                    select_window.destroy()
            
            ttk.Button(select_window, text="Select", command=select_asset).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load assets: {str(e)}")
    
    def remove_asset(self):
        """Remove selected asset from list"""
        selection = self.asset_listbox.curselection()
        if selection:
            self.asset_listbox.delete(selection[0])
    
    def create_ticket(self):
        """Create the ticket"""
        # Validate
        if not self.title_entry.get():
            messagebox.showwarning("Validation", "Please enter a title")
            return
        
        if not self.ticket_type.get():
            messagebox.showwarning("Validation", "Please select a ticket type")
            return
        
        try:
            # Load ticket data
            with open(TIC_PATH, 'r') as f:
                data = json.load(f)
            
            # Generate ticket ID
            next_num = data.get('next_ticket_number', 1)
            ticket_id = f"LAT{next_num:06d}"
            
            # Get assets from listbox
            assets = list(self.asset_listbox.get(0, 'end'))
            
            # Create ticket
            new_ticket = {
                "ticket_id": ticket_id,
                "type": self.ticket_type.get(),
                "title": self.title_entry.get(),
                "description": self.description_text.get('1.0', 'end-1c'),
                "assets": assets,
                "status": "Open",
                "created_date": datetime.now().isoformat(),
                "created_by": self.user,
                "closed_date": None,
                "closed_by": None,
                "actions": [],
                "linked_costs": [],
                "attached_files": []
            }
            
            # Create ticket folder
            create_ticket_folder(ticket_id)
            
            # Add to tickets list
            if 'tickets' not in data:
                data['tickets'] = []
            data['tickets'].append(new_ticket)
            data['next_ticket_number'] = next_num + 1
            
            # Save via temp
            save_json(TIC_PATH, data)
            
            self.callback()
            self.window.destroy()
            messagebox.showinfo("Success", f"Ticket {ticket_id} created successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create ticket: {str(e)}")

class EditTicketWindow:
    def __init__(self, parent, user, callback, ticket_data):
        self.user = user
        self.callback = callback
        self.ticket_data = ticket_data
        self.window = tk.Toplevel(parent)
        self.window.title(f"Edit Ticket - {ticket_data.get('ticket_id')}")
        self.window.geometry("800x700")
        self.window.attributes('-topmost', True)
        
        # Title
        ttk.Label(self.window, text=f"Edit Ticket - {ticket_data.get('ticket_id')}", 
                 font=('TkDefaultFont', 14, 'bold')).pack(pady=10)
        
        # Main frame
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # Ticket Type
        ttk.Label(main_frame, text="Ticket Type:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=0, column=0, sticky='w', pady=5)
        self.ticket_type = ttk.Combobox(main_frame, values=TICKET_TYPES, 
                                        state='readonly', width=30)
        self.ticket_type.grid(row=0, column=1, sticky='ew', pady=5)
        self.ticket_type.set(ticket_data.get('type', 'General'))
        
        # Title
        ttk.Label(main_frame, text="Title:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=1, column=0, sticky='w', pady=5)
        self.title_entry = ttk.Entry(main_frame, width=50)
        self.title_entry.grid(row=1, column=1, sticky='ew', pady=5)
        self.title_entry.insert(0, ticket_data.get('title', ''))
        
        # Description
        ttk.Label(main_frame, text="Description:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=2, column=0, sticky='nw', pady=5)
        self.description_text = tk.Text(main_frame, height=4, width=50)
        self.description_text.grid(row=2, column=1, sticky='ew', pady=5)
        self.description_text.insert('1.0', ticket_data.get('description', ''))
        
        # Assets
        ttk.Label(main_frame, text="Assets:", font=('TkDefaultFont', 10, 'bold')).grid(
            row=3, column=0, sticky='nw', pady=5)
        
        asset_frame = ttk.Frame(main_frame)
        asset_frame.grid(row=3, column=1, sticky='ew', pady=5)
        
        # Asset listbox
        self.asset_listbox = tk.Listbox(asset_frame, height=5)
        self.asset_listbox.pack(side='left', fill='both', expand=True)
        
        asset_scroll = ttk.Scrollbar(asset_frame, orient='vertical', 
                                     command=self.asset_listbox.yview)
        asset_scroll.pack(side='right', fill='y')
        self.asset_listbox.config(yscrollcommand=asset_scroll.set)
        
        # Load existing assets
        assets = ticket_data.get('assets', [])
        if not isinstance(assets, list):
            assets = [ticket_data.get('asset', '')]
        for asset in assets:
            if asset:
                self.asset_listbox.insert('end', asset)
        
        # Asset buttons
        asset_btn_frame = ttk.Frame(main_frame)
        asset_btn_frame.grid(row=4, column=1, sticky='w', pady=5)
        ttk.Button(asset_btn_frame, text="Add Asset", command=self.add_asset).pack(side='left', padx=2)
        ttk.Button(asset_btn_frame, text="Remove Selected", command=self.remove_asset).pack(side='left', padx=2)
        
        main_frame.grid_columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Save Changes", command=self.save_ticket).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def add_asset(self):
        """Add an asset to the ticket"""
        try:
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            
            asset_ids = [item.get('identifier', '') for item in index_data if item.get('identifier')]
            
            if not asset_ids:
                messagebox.showwarning("No Assets", "No assets found in index.")
                return
            
            select_window = tk.Toplevel(self.window)
            select_window.title("Select Asset")
            select_window.geometry("400x500")
            select_window.attributes('-topmost', True)
            
            ttk.Label(select_window, text="Select Asset:", font=('TkDefaultFont', 10, 'bold')).pack(pady=10)
            
            search_var = tk.StringVar()
            search_entry = ttk.Entry(select_window, textvariable=search_var)
            search_entry.pack(fill='x', padx=10, pady=5)
            
            listbox_frame = ttk.Frame(select_window)
            listbox_frame.pack(fill='both', expand=True, padx=10, pady=5)
            
            asset_list = tk.Listbox(listbox_frame)
            asset_list.pack(side='left', fill='both', expand=True)
            
            scroll = ttk.Scrollbar(listbox_frame, orient='vertical', command=asset_list.yview)
            scroll.pack(side='right', fill='y')
            asset_list.config(yscrollcommand=scroll.set)
            
            def filter_assets(*args):
                asset_list.delete(0, 'end')
                search_term = search_var.get().lower()
                for asset_id in sorted(asset_ids):
                    if search_term in asset_id.lower():
                        asset_list.insert('end', asset_id)
            
            search_var.trace('w', filter_assets)
            filter_assets()
            
            def select_asset():
                selection = asset_list.curselection()
                if selection:
                    selected = asset_list.get(selection[0])
                    if selected not in self.asset_listbox.get(0, 'end'):
                        self.asset_listbox.insert('end', selected)
                    select_window.destroy()
            
            ttk.Button(select_window, text="Select", command=select_asset).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load assets: {str(e)}")
    
    def remove_asset(self):
        """Remove selected asset from list"""
        selection = self.asset_listbox.curselection()
        if selection:
            self.asset_listbox.delete(selection[0])
    
    def save_ticket(self):
        """Save changes to the ticket"""
        if not self.title_entry.get():
            messagebox.showwarning("Validation", "Please enter a title")
            return
        
        if not self.ticket_type.get():
            messagebox.showwarning("Validation", "Please select a ticket type")
            return
        
        try:
            with open(TIC_PATH, 'r') as f:
                data = json.load(f)
            
            assets = list(self.asset_listbox.get(0, 'end'))
            
            for t in data.get('tickets', []):
                if t.get('ticket_id') == self.ticket_data.get('ticket_id'):
                    t['type'] = self.ticket_type.get()
                    t['title'] = self.title_entry.get()
                    t['description'] = self.description_text.get('1.0', 'end-1c')
                    t['assets'] = assets
                    break
            
            save_json(TIC_PATH, data)
            
            self.callback()
            self.window.destroy()
            messagebox.showinfo("Success", f"Ticket {self.ticket_data.get('ticket_id')} updated successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save ticket: {str(e)}")

class AddActionWindow:
    def __init__(self, parent, ticket, user, callback, action_to_edit=None):
        self.ticket = ticket
        self.callback = callback
        self.user = user
        self.action_to_edit = action_to_edit
        self.window = tk.Toplevel(parent)
        
        if action_to_edit:
            self.window.title(f"Edit Action - {action_to_edit.get('action_id')}")
        else:
            self.window.title(f"Add Action to {ticket.get('ticket_id')}")
        
        self.window.geometry("700x500")
        self.window.attributes('-topmost', True)
        
        # Action Type
        ttk.Label(self.window, text="Action Type:", font=('TkDefaultFont', 10)).grid(row=0, column=0, sticky='w', padx=10, pady=10)
        self.action_type = ttk.Combobox(self.window, values=ASSET_SETTINGS.get('ticket_types', []), state='readonly', width=30)
        self.action_type.grid(row=0, column=1, padx=10, pady=10)
        
        if action_to_edit:
            self.action_type.set(action_to_edit.get('type', ''))
        else:
            self.action_type.current(0)
        
        self.action_type.bind('<<ComboboxSelected>>', self.on_type_change)
        
        # User
        ttk.Label(self.window, text="User:", font=('TkDefaultFont', 10)).grid(row=1, column=0, sticky='w', padx=10, pady=10)
        self.user_var = tk.StringVar(value=action_to_edit.get('user', user) if action_to_edit else user)
        self.user_dropdown = ttk.Combobox(self.window, textvariable=self.user_var, 
                                          values=ASSET_SETTINGS.get('users', []), state='readonly', width=30)
        self.user_dropdown.grid(row=1, column=1, padx=10, pady=10)
        
        # Description
        ttk.Label(self.window, text="Description:", font=('TkDefaultFont', 10)).grid(row=2, column=0, sticky='nw', padx=10, pady=10)
        self.description = tk.Text(self.window, width=32, height=8)
        self.description.grid(row=2, column=1, padx=10, pady=10)
        
        if action_to_edit:
            self.description.insert('1.0', action_to_edit.get('description', ''))
        
        # Referenced Assets
        ttk.Label(self.window, text="Referenced Assets:", font=('TkDefaultFont', 10)).grid(row=3, column=0, sticky='nw', padx=10, pady=10)
        
        ref_assets_frame = ttk.Frame(self.window)
        ref_assets_frame.grid(row=3, column=1, padx=10, pady=10, sticky='ew')
        
        # Get all assets in ticket
        ticket_assets = self.ticket.get('assets', [])
        if not isinstance(ticket_assets, list):
            ticket_assets = [self.ticket.get('asset', '')]
        ticket_assets = [a for a in ticket_assets if a]  # Filter empty
        
        # Create checkboxes for each asset
        self.asset_check_vars = {}
        for i, asset_id in enumerate(ticket_assets):
            var = tk.BooleanVar()
            # Check if this asset was previously selected
            if action_to_edit:
                referenced = action_to_edit.get('referenced_assets', [])
                var.set(asset_id in referenced)
            else:
                var.set(True)  # Default to all selected
            
            cb = ttk.Checkbutton(ref_assets_frame, text=asset_id, variable=var)
            cb.grid(row=i, column=0, sticky='w', pady=2)
            self.asset_check_vars[asset_id] = var
        
        # File attachment section
        ttk.Label(self.window, text="Files:", font=('TkDefaultFont', 10)).grid(row=4, column=0, sticky='nw', padx=10, pady=10)
        
        # Frame for file list and buttons
        file_container = ttk.Frame(self.window)
        file_container.grid(row=4, column=1, padx=10, pady=10, sticky='ew')
        
        # Listbox to show attached files
        self.files_listbox = tk.Listbox(file_container, height=3)
        self.files_listbox.pack(fill='both', expand=True)
        
        # Store file info
        self.attached_files = []
        if action_to_edit and action_to_edit.get('files'):
            for file_info in action_to_edit['files']:
                self.attached_files.append(file_info)
                self.files_listbox.insert('end', file_info.get('filename', ''))
        elif action_to_edit and action_to_edit.get('file'):
            # Handle old single-file format
            self.attached_files.append(action_to_edit['file'])
            self.files_listbox.insert('end', action_to_edit['file'].get('filename', ''))
        
        # File buttons
        file_btn_frame = ttk.Frame(file_container)
        file_btn_frame.pack(fill='x', pady=5)
        ttk.Button(file_btn_frame, text="Add File", command=self.add_file).pack(side='left', padx=2)
        ttk.Button(file_btn_frame, text="Open Selected", command=self.open_file).pack(side='left', padx=2)
        ttk.Button(file_btn_frame, text="Remove Selected", command=self.remove_file).pack(side='left', padx=2)
        
        self.file_required_label = ttk.Label(self.window, text="", foreground="red")
        self.file_required_label.grid(row=5, column=1, sticky='w', padx=10)
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        save_text = "Update" if action_to_edit else "Add"
        ttk.Button(btn_frame, text=save_text, command=self.add_action).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
        
        self.on_type_change(None)
    
    def add_file(self):
        """Add a file to the action"""
        file_path = filedialog.askopenfilename(title="Select file to attach")
        if file_path:
            # Copy file to ticket folder
            copied_path = copy_file_to_ticket(file_path, self.ticket.get('ticket_id'))
            
            if not copied_path:
                messagebox.showerror("Error", "Failed to copy file to ticket folder")
                return
            
            file_info = {
                'filename': Path(copied_path).name,
                'path': copied_path,
                'uploaded_by': self.user_var.get(),
                'uploaded_date': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            }
            self.attached_files.append(file_info)
            self.files_listbox.insert('end', file_info['filename'])
    
    def open_file(self):
        """Open the selected file"""
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a file to open")
            return
        
        file_info = self.attached_files[selection[0]]
        file_path = file_info.get('path', '')
        
        if file_path and Path(file_path).exists():
            try:
                os.startfile(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {str(e)}")
        else:
            messagebox.showerror("Error", "File not found")
    
    def remove_file(self):
        """Remove the selected file from the list"""
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a file to remove")
            return
        
        if messagebox.askyesno("Confirm", "Delete this file from the ticket folder as well?"):
            try:
                file_info = self.attached_files[idx]
                file_path = file_info.get('path', '')
                if file_path and Path(file_path).exists():
                    Path(file_path).unlink()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete file: {str(e)}")

        idx = selection[0]
        self.attached_files.pop(idx)
        self.files_listbox.delete(idx)
    
    def on_type_change(self, event):
        action_type = self.action_type.get()
        if action_type in ['Dispatch', 'Return']:
            self.file_required_label.config(text=f"* At least one file required for {action_type} actions")
        else:
            self.file_required_label.config(text="")
    
    def add_action(self):
        action_type = self.action_type.get()
        
        # Check if file is required
        if action_type in ['Dispatch', 'Return'] and len(self.attached_files) == 0:
            messagebox.showwarning("Warning", f"At least one file attachment is required for {action_type} actions")
            return
        
        # Get selected assets
        referenced_assets = [asset_id for asset_id, var in self.asset_check_vars.items() if var.get()]
        
        try:
            # Load ticket data
            with open(TIC_PATH, 'r') as f:
                ticket_data = json.load(f)
            
            for t in ticket_data.get('tickets', []):
                if t.get('ticket_id') == self.ticket.get('ticket_id'):
                    actions = t.get('actions', [])
                    
                    if self.action_to_edit:
                        # Update existing action
                        for action in actions:
                            if action.get('action_id') == self.action_to_edit.get('action_id'):
                                action['type'] = self.action_type.get()
                                action['user'] = self.user_var.get()
                                action['description'] = self.description.get('1.0', 'end-1c')
                                action['referenced_assets'] = referenced_assets
                                
                                # Update files
                                if self.attached_files:
                                    action['files'] = self.attached_files
                                    
                                    # Also update ticket's attached_files
                                    attached_files = t.get('attached_files', [])
                                    for file_info in self.attached_files:
                                        # Only add if not already in list
                                        if not any(f.get('filename') == file_info.get('filename') for f in attached_files):
                                            attached_files.append(file_info)
                                    t['attached_files'] = attached_files
                                break
                    else:
                        # Add new action
                        action_num = len(actions) + 1
                        action_id = f"{self.ticket.get('ticket_id')}-{action_num}"
                        
                        new_action = {
                            'action_id': action_id,
                            'type': self.action_type.get(),
                            'user': self.user_var.get(),
                            'date': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                            'description': self.description.get('1.0', 'end-1c'),
                            'referenced_assets': referenced_assets
                        }
                        
                        # Add files if attached
                        if self.attached_files:
                            new_action['files'] = self.attached_files
                            
                            # Also add to ticket's attached_files
                            attached_files = t.get('attached_files', [])
                            for file_info in self.attached_files:
                                attached_files.append(file_info)
                            t['attached_files'] = attached_files
                        
                        actions.append(new_action)
                        
                        # Handle Dispatch/Return status changes - only for referenced assets
                        if self.action_type.get() == 'Dispatch':
                            self.update_asset_status(referenced_assets, 'External')
                        elif self.action_type.get() == 'Return':
                            self.update_asset_status(referenced_assets, 'Available')
                    
                    t['actions'] = actions
                    break
            
            # Save with backup
            save_json(TIC_PATH, ticket_data)
            
            self.callback()
            self.window.destroy()
            
            success_msg = "Action updated successfully" if self.action_to_edit else "Action added successfully"
            messagebox.showinfo("Success", success_msg)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to {'update' if self.action_to_edit else 'add'} action: {str(e)}")
    
    def update_asset_status(self, asset_ids, new_status):
        """Update asset status in registry"""
        try:
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            if not isinstance(asset_ids, list):
                asset_ids = [asset_ids]
            
            for asset_id in asset_ids:
                for reg_item in registry_data:
                    if reg_item.get('identifier') == asset_id:
                        reg_item['status'] = new_status
                        break
            
            # Save with backup
            save_json(REG_PATH, registry_data)
        except Exception as e:
            print(f"Failed to update asset status: {str(e)}")

class AddCostWindow:
    def __init__(self, parent, ticket, user, callback, cost_to_edit=None):
        self.ticket = ticket
        self.callback = callback
        self.cost_to_edit = cost_to_edit
        self.window = tk.Toplevel(parent)
        
        if cost_to_edit:
            self.window.title(f"Edit Cost - {ticket.get('ticket_id')}")
        else:
            self.window.title(f"Add Cost to {ticket.get('ticket_id')}")
        self.window.geometry("900x700")
        self.window.attributes('-topmost', True)
        
        # Create scrollable frame
        canvas = tk.Canvas(self.window)
        scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Supplier Info - Two columns
        # Left Column
        ttk.Label(self.scrollable_frame, text="Supplier:", font=('TkDefaultFont', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.supplier = ttk.Entry(self.scrollable_frame, width=30)
        self.supplier.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.supplier.insert(0, cost_to_edit.get('supplier', ''))
        
        ttk.Label(self.scrollable_frame, text="Address:").grid(row=1, column=0, sticky='w', padx=10, pady=5)
        self.supplier_address = ttk.Entry(self.scrollable_frame, width=30)
        self.supplier_address.grid(row=1, column=1, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.supplier_address.insert(0, cost_to_edit.get('supplier_address', ''))
        
        ttk.Label(self.scrollable_frame, text="Email:").grid(row=2, column=0, sticky='w', padx=10, pady=5)
        self.supplier_email = ttk.Entry(self.scrollable_frame, width=30)
        self.supplier_email.grid(row=2, column=1, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.supplier_email.insert(0, cost_to_edit.get('supplier_email', ''))
        
        ttk.Label(self.scrollable_frame, text="Contact Number:").grid(row=3, column=0, sticky='w', padx=10, pady=5)
        self.supplier_contact = ttk.Entry(self.scrollable_frame, width=30)
        self.supplier_contact.grid(row=3, column=1, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.supplier_contact.insert(0, cost_to_edit.get('supplier_contact', ''))
        
        ttk.Label(self.scrollable_frame, text="Currency:").grid(row=4, column=0, sticky='w', padx=10, pady=5)
        # Handle both old and new currency format
        currencies = ASSET_SETTINGS.get('currencies', [{'code': 'GBP', 'symbol': '£'}])
        currency_codes = [c['code'] if isinstance(c, dict) else c for c in currencies]
        self.currency = ttk.Combobox(self.scrollable_frame, values=currency_codes, state='readonly', width=28)
        self.currency.grid(row=4, column=1, padx=10, pady=5, sticky='w')
        if cost_to_edit and cost_to_edit.get('currency'):
            self.currency.set(cost_to_edit.get('currency'))
        else:
            self.currency.current(0)
        
        # Right Column
        ttk.Label(self.scrollable_frame, text="Delivery Address:").grid(row=0, column=2, sticky='w', padx=10, pady=5)
        self.delivery_address = ttk.Combobox(self.scrollable_frame, values=ASSET_SETTINGS.get('delivery_addresses', []), width=28)
        self.delivery_address.grid(row=0, column=3, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.delivery_address.set(cost_to_edit.get('delivery_address', ''))
        
        ttk.Label(self.scrollable_frame, text="Date:").grid(row=1, column=2, sticky='w', padx=10, pady=5)
        self.date = ttk.Entry(self.scrollable_frame, width=30)
        if cost_to_edit:
            self.date.insert(0, cost_to_edit.get('date', ''))
        else:
            self.date.insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.date.grid(row=1, column=3, padx=10, pady=5, sticky='w')
        
        ttk.Label(self.scrollable_frame, text="PO Number:").grid(row=2, column=2, sticky='w', padx=10, pady=5)
        self.po_number = ttk.Entry(self.scrollable_frame, width=30)
        self.po_number.grid(row=2, column=3, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.po_number.insert(0, cost_to_edit.get('po_number', ''))
        
        ttk.Label(self.scrollable_frame, text="Signoff:").grid(row=3, column=2, sticky='w', padx=10, pady=5)
        self.signoff = ttk.Entry(self.scrollable_frame, width=30)
        self.signoff.grid(row=3, column=3, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.signoff.insert(0, cost_to_edit.get('signoff', ''))

        ttk.Label(self.scrollable_frame, text="RMA:").grid(row=5, column=2, sticky='w', padx=10, pady=5)
        self.rma = ttk.Entry(self.scrollable_frame, width=30)
        self.rma.grid(row=5, column=3, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.rma.insert(0, cost_to_edit.get('rma', ''))
        
        ttk.Label(self.scrollable_frame, text="Carriage:").grid(row=4, column=2, sticky='w', padx=10, pady=5)
        self.carriage = ttk.Entry(self.scrollable_frame, width=30)
        if cost_to_edit:
            self.carriage.insert(0, str(cost_to_edit.get('carriage', '0')))
        else:
            self.carriage.insert(0, "0")
        self.carriage.grid(row=4, column=3, padx=10, pady=5, sticky='w')
        self.carriage.bind('<KeyRelease>', lambda e: self.calculate_final_cost())
        
        # Instructions - Full width single line
        ttk.Label(self.scrollable_frame, text="Instructions:").grid(row=6, column=0, sticky='w', padx=10, pady=5)
        self.instructions = ttk.Entry(self.scrollable_frame, width=95)
        self.instructions.grid(row=6, column=1, columnspan=3, padx=10, pady=5, sticky='w')
        if cost_to_edit:
            self.instructions.insert(0, cost_to_edit.get('instructions', ''))
        
        # Line Items
        ttk.Label(self.scrollable_frame, text="Line Items:", font=('TkDefaultFont', 10, 'bold')).grid(row=7, column=0, columnspan=4, sticky='w', padx=10, pady=(20, 5))
        
        self.line_items = []
        self.line_frame = ttk.Frame(self.scrollable_frame)
        self.line_frame.grid(row=8, column=0, columnspan=4, padx=10, pady=5)
        
        # Add headers first
        ttk.Label(self.line_frame, text="Qty", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=0, padx=2, pady=2)
        ttk.Label(self.line_frame, text="Description", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=1, padx=2, pady=2)
        ttk.Label(self.line_frame, text="Cost", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=2, padx=2, pady=2)
        ttk.Label(self.line_frame, text="Discount%", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=3, padx=2, pady=2)
        ttk.Label(self.line_frame, text="Customer", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=4, padx=2, pady=2)
        ttk.Label(self.line_frame, text="Reference", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=5, padx=2, pady=2)
        ttk.Label(self.line_frame, text="GL Code", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=6, padx=2, pady=2)
        ttk.Label(self.line_frame, text="Actions", font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=7, padx=2, pady=2)
        
        ttk.Button(self.scrollable_frame, text="Add Line Item", command=self.add_line_item).grid(row=9, column=0, columnspan=4, pady=10)
        
        # Final Cost (Auto-calculated)
        ttk.Label(self.scrollable_frame, text="Final Cost:", font=('TkDefaultFont', 10, 'bold')).grid(row=10, column=0, sticky='w', padx=10, pady=5)
        self.final_cost_label = ttk.Label(self.scrollable_frame, text="0.00", font=('TkDefaultFont', 10))
        self.final_cost_label.grid(row=10, column=1, padx=10, pady=5, sticky='w')
        
        # Buttons
        btn_frame = ttk.Frame(self.scrollable_frame)
        btn_frame.grid(row=11, column=0, columnspan=4, pady=20)
        ttk.Button(btn_frame, text="Save", command=self.save_cost).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Generate Purchase Requisition", command=self.generate_pr).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add initial line item or populate from cost_to_edit
        if cost_to_edit and cost_to_edit.get('line_items'):
            for line_item in cost_to_edit.get('line_items', []):
                self.add_line_item(line_item)
        else:
            self.add_line_item()
        
        # Recalculate if editing
        if cost_to_edit:
            self.calculate_final_cost()
    
    def calculate_final_cost(self):
        """Calculate final cost from line items and carriage"""
        try:
            total = 0
            for line in self.line_items:
                qty = float(line['qty'].get() or 0)
                cost = float(line['cost'].get() or 0)
                discount = float(line['discount'].get() or 0)
                total += qty * cost * (1 - discount / 100)
            
            carriage = float(self.carriage.get() or 0)
            final = total + carriage
            
            self.final_cost_label.config(text=f"{final:.2f}")
        except:
            self.final_cost_label.config(text="Error")
    
    def add_line_item(self, line_data=None):
        row = len(self.line_items) + 1  # +1 to account for header row
        
        line = {}
        line['qty'] = ttk.Entry(self.line_frame, width=8)
        line['qty'].grid(row=row, column=0, padx=2, pady=2)
        line['qty'].bind('<KeyRelease>', lambda e: self.calculate_final_cost())
        if line_data:
            line['qty'].insert(0, str(line_data.get('quantity', '')))
        
        line['desc'] = ttk.Entry(self.line_frame, width=25)
        line['desc'].grid(row=row, column=1, padx=2, pady=2)
        if line_data:
            line['desc'].insert(0, line_data.get('description', ''))
        
        line['cost'] = ttk.Entry(self.line_frame, width=10)
        line['cost'].grid(row=row, column=2, padx=2, pady=2)
        line['cost'].bind('<KeyRelease>', lambda e: self.calculate_final_cost())
        if line_data:
            line['cost'].insert(0, str(line_data.get('cost', '')))
        
        line['discount'] = ttk.Entry(self.line_frame, width=8)
        line['discount'].grid(row=row, column=3, padx=2, pady=2)
        line['discount'].bind('<KeyRelease>', lambda e: self.calculate_final_cost())
        if line_data:
            line['discount'].insert(0, str(line_data.get('discount', '')))
        
        line['customer'] = ttk.Entry(self.line_frame, width=15)
        line['customer'].grid(row=row, column=4, padx=2, pady=2)
        if line_data:
            line['customer'].insert(0, line_data.get('customer', ''))
        
        line['reference'] = ttk.Entry(self.line_frame, width=12)
        line['reference'].grid(row=row, column=5, padx=2, pady=2)
        if line_data:
            line['reference'].insert(0, line_data.get('reference', ''))
        
        line['gl_code'] = ttk.Combobox(self.line_frame, values=ASSET_SETTINGS.get('gl_codes', []), width=20)
        line['gl_code'].grid(row=row, column=6, padx=2, pady=2)
        if line_data:
            line['gl_code'].set(line_data.get('gl_code', ''))
        
        # Action buttons frame
        action_frame = ttk.Frame(self.line_frame)
        action_frame.grid(row=row, column=7, padx=2, pady=2)
        
        up_btn = ttk.Button(action_frame, text="↑", width=2, command=lambda: self.move_line_up(line))
        up_btn.pack(side='left', padx=1)
        line['up_btn'] = up_btn
        
        down_btn = ttk.Button(action_frame, text="↓", width=2, command=lambda: self.move_line_down(line))
        down_btn.pack(side='left', padx=1)
        line['down_btn'] = down_btn
        
        remove_btn = ttk.Button(action_frame, text="X", width=2, command=lambda: self.remove_line_item(line))
        remove_btn.pack(side='left', padx=1)
        line['remove_btn'] = remove_btn
        
        line['action_frame'] = action_frame
        
        self.line_items.append(line)
        self.calculate_final_cost()
    
    def remove_line_item(self, line):
        if len(self.line_items) > 1:
            self.line_items.remove(line)
            for widget in line.values():
                if isinstance(widget, (ttk.Entry, ttk.Combobox, ttk.Button, ttk.Frame)):
                    widget.destroy()
            self.refresh_line_items()
            self.calculate_final_cost()
    
    def move_line_up(self, line):
        idx = self.line_items.index(line)
        if idx > 0:
            self.line_items[idx], self.line_items[idx - 1] = self.line_items[idx - 1], self.line_items[idx]
            self.refresh_line_items()
    
    def move_line_down(self, line):
        idx = self.line_items.index(line)
        if idx < len(self.line_items) - 1:
            self.line_items[idx], self.line_items[idx + 1] = self.line_items[idx + 1], self.line_items[idx]
            self.refresh_line_items()
    
    def refresh_line_items(self):
        """Refresh the display of line items after reordering or removal"""
        for idx, line in enumerate(self.line_items):
            row = idx + 1  # +1 for header row
            line['qty'].grid(row=row, column=0)
            line['desc'].grid(row=row, column=1)
            line['cost'].grid(row=row, column=2)
            line['discount'].grid(row=row, column=3)
            line['customer'].grid(row=row, column=4)
            line['reference'].grid(row=row, column=5)
            line['gl_code'].grid(row=row, column=6)
            line['action_frame'].grid(row=row, column=7)
    
    def save_cost(self):
        try:
            lines = []
            total_lines = 0
            for line in self.line_items:
                qty = float(line['qty'].get() or 0)
                cost = float(line['cost'].get() or 0)
                discount = float(line['discount'].get() or 0)
                total = qty * cost * (1 - discount / 100)
                total_lines += total
                
                lines.append({
                    'quantity': qty,
                    'description': line['desc'].get(),
                    'cost': cost,
                    'discount': discount,
                    'total_cost': total,
                    'customer': line['customer'].get(),
                    'reference': line['reference'].get(),
                    'gl_code': line['gl_code'].get()
                })
            
            carriage = float(self.carriage.get() or 0)
            final_cost = total_lines + carriage
            rma = self.rma.get() if hasattr(self, 'rma') else ''
            
            # Load costs data
            if COST_PATH.exists():
                with open(COST_PATH, 'r') as f:
                    cost_data = json.load(f)
            else:
                cost_data = {"next_cost_number": 1, "costs": []}
            
            if self.cost_to_edit:
                # Update existing cost
                cost_id = self.cost_to_edit.get('cost_id')
                for i, c in enumerate(cost_data.get('costs', [])):
                    if c.get('cost_id') == cost_id:
                        cost_data['costs'][i] = {
                            'cost_id': cost_id,
                            'ticket_id': self.ticket.get('ticket_id'),
                            'supplier': self.supplier.get(),
                            'supplier_address': self.supplier_address.get(),
                            'supplier_email': self.supplier_email.get(),
                            'supplier_contact': self.supplier_contact.get(),
                            'currency': self.currency.get(),
                            'delivery_address': self.delivery_address.get(),
                            'instructions': self.instructions.get(),
                            'po_number': self.po_number.get(),
                            'date': self.date.get(),
                            'signoff': self.signoff.get(),
                            'carriage': carriage,
                            'final_cost': final_cost,
                            'rma': rma,
                            'line_items': lines
                        }
                        break
            else:
                # Add new cost with unique ID
                next_num = cost_data.get('next_cost_number', 1)
                cost_id = f"COST{next_num:06d}"
                
                new_cost = {
                    'cost_id': cost_id,
                    'ticket_id': self.ticket.get('ticket_id'),
                    'supplier': self.supplier.get(),
                    'supplier_address': self.supplier_address.get(),
                    'supplier_email': self.supplier_email.get(),
                    'supplier_contact': self.supplier_contact.get(),
                    'currency': self.currency.get(),
                    'delivery_address': self.delivery_address.get(),
                    'instructions': self.instructions.get(),
                    'po_number': self.po_number.get(),
                    'date': self.date.get(),
                    'signoff': self.signoff.get(),
                    'carriage': carriage,
                    'final_cost': final_cost,
                    'rma': rma,
                    'line_items': lines
                }
                
                if 'costs' not in cost_data:
                    cost_data['costs'] = []
                cost_data['costs'].append(new_cost)
                cost_data['next_cost_number'] = next_num + 1
            
            # Save with backup
            save_json(COST_PATH, cost_data)
            
            # Add action to ticket
            with open(TIC_PATH, 'r') as f:
                ticket_data = json.load(f)
            
            for t in ticket_data.get('tickets', []):
                if t.get('ticket_id') == self.ticket.get('ticket_id'):
                    actions = t.get('actions', [])
                    action_num = len(actions) + 1
                    action_id = f"{self.ticket.get('ticket_id')}-{action_num}"
                    
                    if self.cost_to_edit:
                        description = f"Updated cost {cost_id} - {self.supplier.get()}"
                    else:
                        description = f"Added cost {cost_id} - {self.supplier.get()}"
                    
                    actions.append({
                        'action_id': action_id,
                        'type': 'Cost Entry',
                        'user': self.ticket.get('created_by', 'System'),
                        'date': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                        'description': description
                    })
                    t['actions'] = actions
                    break
            
            # Save ticket data to temp then main
            save_json(TIC_PATH, ticket_data)
            
            self.callback()
            self.window.destroy()
            
            success_msg = "Cost updated successfully" if self.cost_to_edit else "Cost added successfully"
            messagebox.showinfo("Success", success_msg)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save cost: {str(e)}")
    
    def generate_pr(self):
        """Generate Purchase Requisition from cost data"""
        try:
            # Check if template exists
            template_path = Path(__file__).parent / "assets" / "pr_template.xlsx"
            if not template_path.exists():
                messagebox.showerror("Error", "Purchase Requisition template not found at assets/pr_template.xlsx")
                return
            
            # Get ticket folder
            ticket_folder = create_ticket_folder(self.ticket.get('ticket_id'))
            
            # Create output filename
            output_filename = f"{self.ticket.get('ticket_id')} Req.xlsx"
            output_path = ticket_folder / output_filename
            
            # Copy template to ticket folder
            shutil.copy2(template_path, output_path)
            
            # Load the workbook
            wb = load_workbook(output_path)
            ws = wb.active
            
            # Get currency symbol from settings - use hardcoded symbols to avoid encoding issues
            selected_currency = self.currency.get()
            # Map currency codes to symbols directly to avoid encoding issues
            currency_map = {
                'GBP': '\u00A3',  # £
                'EUR': '\u20AC',  # €
                'USD': '$'
            }
            currency_symbol = currency_map.get(selected_currency, '\u00A3')
            
            # Fill in supplier information
            # Format date as DD/MM/YYYY
            date_str = self.date.get()
            try:
                # Try to parse and reformat if needed
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                ws['P3'] = date_obj.strftime('%d/%m/%Y')
            except:
                # If already in correct format or other format, use as is
                ws['P3'] = date_str
            
            # Format supplier address with newlines
            address = self.supplier_address.get()
            if address:
                # Replace commas with newlines for address formatting
                formatted_address = address.replace(',', '\n')
                ws['A7'] = f"{self.supplier.get()}\n{formatted_address}"
            else:
                ws['A7'] = self.supplier.get()
            
            ws['C11'] = self.supplier_contact.get()
            ws['C12'] = self.supplier_email.get()
            ws['E38'] = self.instructions.get()
            
            # Format carriage with currency symbol
            try:
                carriage_val = float(self.carriage.get() or 0)
                ws['I34'] = f"{carriage_val:.2f}"
            except:
                ws['I34'] = self.carriage.get()
            
            ws['E40'] = self.signoff.get()
            
            # Add line items
            line_rows = [16, 18, 20, 22, 24, 26, 28, 30, 32]
            
            for idx, line in enumerate(self.line_items):
                if idx >= 9:  # Maximum 9 lines
                    break
                
                row = line_rows[idx]
                ws[f'A{row}'] = line['qty'].get()
                ws[f'D{row}'] = line['desc'].get()
                
                # Format cost with currency symbol
                try:
                    cost_val = float(line['cost'].get() or 0)
                    ws[f'G{row}'] = f"{currency_symbol}{cost_val:.2f}"
                except:
                    ws[f'G{row}'] = line['cost'].get()
                
                # Format discount with % symbol
                discount_str = line['discount'].get()
                if discount_str:
                    ws[f'H{row}'] = f"{discount_str}%"
                

                ref = line['reference'].get()
                gl = line['gl_code'].get()

                if ref and gl:
                    ws[f'Q{row}'] = f"{ref}, Capex {gl}"
                elif ref:
                    ws[f'Q{row}'] = ref
                elif gl:
                    ws[f'Q{row}'] = f"Capex {gl}"
                ws[f'N{row}'] = line['customer'].get()
            
            # Save the workbook
            wb.save(output_path)
            wb.close()
            
            # Attach the generated file to the ticket
            try:
                with open(TIC_PATH, 'r') as f:
                    ticket_data = json.load(f)
                
                for t in ticket_data.get('tickets', []):
                    if t.get('ticket_id') == self.ticket.get('ticket_id'):
                        # Add to attached files
                        attached_files = t.get('attached_files', [])
                        attached_files.append({
                            'filename': output_filename,
                            'path': str(output_path),
                            'uploaded_by': self.ticket.get('created_by', 'System'),
                            'uploaded_date': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                        })
                        t['attached_files'] = attached_files
                        break
                
                # Save with backup
                save_json(TIC_PATH, ticket_data)
                
            except Exception as e:
                print(f"Warning: Failed to attach PR to ticket: {str(e)}")
            
            messagebox.showinfo("Success", f"Purchase Requisition generated and attached to ticket:\n{output_path}")
            
            # Ask if they want to open it
            if messagebox.askyesno("Open File", "Would you like to open the generated Purchase Requisition?"):
                try:
                    os.startfile(str(output_path))
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to open file: {str(e)}")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Purchase Requisition: {str(e)}")

class CostDetailWindow:
    def __init__(self, parent, ticket, cost, callback):
        self.ticket = ticket
        self.cost = cost
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title(f"Cost Detail - {cost.get('supplier')}")
        self.window.geometry("900x600")
        self.window.attributes('-topmost', True)
        
        # Display cost details
        detail_frame = ttk.Frame(self.window)
        detail_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        row = 0
        for key, value in cost.items():
            if key != 'line_items':
                ttk.Label(detail_frame, text=f"{key.replace('_', ' ').title()}:", font=('TkDefaultFont', 10, 'bold')).grid(row=row, column=0, sticky='w', padx=5, pady=2)
                ttk.Label(detail_frame, text=str(value), wraplength=400).grid(row=row, column=1, sticky='w', padx=5, pady=2)
                row += 1
        
        # Line items
        ttk.Label(detail_frame, text="Line Items:", font=('TkDefaultFont', 11, 'bold')).grid(row=row, column=0, columnspan=2, sticky='w', padx=5, pady=(10, 5))
        row += 1
        
        tree_frame = ttk.Frame(detail_frame)
        tree_frame.grid(row=row, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        tree = ttk.Treeview(tree_frame, yscrollcommand=scrollbar.set)
        scrollbar.config(command=tree.yview)
        
        scrollbar.pack(side='right', fill='y')
        tree.pack(fill='both', expand=True)
        
        tree['columns'] = ('qty', 'desc', 'cost', 'discount', 'total', 'customer', 'ref', 'gl')
        tree.column('#0', width=0, stretch=False)
        
        headers = ['Qty', 'Description', 'Cost', 'Discount', 'Total', 'Customer', 'Reference', 'GL Code']
        for i, col in enumerate(tree['columns']):
            tree.column(col, width=100)
            tree.heading(col, text=headers[i])
        
        for line in cost.get('line_items', []):
            tree.insert('', 'end', values=(
                line.get('quantity', ''),
                line.get('description', ''),
                line.get('cost', ''),
                line.get('discount', ''),
                line.get('total_cost', ''),
                line.get('customer', ''),
                line.get('reference', ''),
                line.get('gl_code', '')
            ))
        
        detail_frame.grid_rowconfigure(row, weight=1)
        detail_frame.grid_columnconfigure(1, weight=1)
        
        # Close button
        ttk.Button(self.window, text="Close", command=self.window.destroy).pack(pady=10)

class EditRegistryWindow:
    def __init__(self, parent, item, callback):
        self.item = item
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title(f"Edit Registry - {item.get('identifier')}")
        self.window.geometry("600x500")
        self.window.attributes('-topmost', True)
        
        self.entries = {}
        row = 0
        
        for key, value in item.items():
            ttk.Label(self.window, text=f"{key.replace('_', ' ').title()}:", font=('TkDefaultFont', 10)).grid(row=row, column=0, sticky='w', padx=10, pady=5)
            
            if key == 'identifier':
                # Make identifier read-only
                entry = ttk.Label(self.window, text=str(value), width=32)
            elif key == 'location':
                entry = ttk.Combobox(self.window, values=ASSET_SETTINGS.get('locations', []), width=30)
                entry.set(str(value) if value else '')
            elif key == 'cal_by':
                entry = ttk.Entry(self.window, width=32)
                entry.insert(0, str(value) if value else '')
            elif key == 'status':
                entry = ttk.Combobox(self.window, values=ASSET_SETTINGS.get('asset_statuses', []), width=30)
                entry.set(str(value) if value else '')
            else:
                entry = ttk.Entry(self.window, width=32)
                entry.insert(0, str(value) if value else '')
            
            entry.grid(row=row, column=1, padx=10, pady=5)
            self.entries[key] = entry
            row += 1
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def save(self):
        try:
            # Load registry
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            for i, reg_item in enumerate(registry_data):
                if reg_item.get('identifier') == self.item.get('identifier'):
                    for key, entry in self.entries.items():
                        if key == 'identifier':
                            continue  # Skip identifier - it's read-only
                        registry_data[i][key] = entry.get()
                    break
            
            # Save with backup
            save_json(REG_PATH, registry_data)
            
            self.callback()
            self.window.destroy()
            messagebox.showinfo("Success", "Registry updated successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {str(e)}")

class ViewAssetTicketsWindow:
    def __init__(self, parent, asset_id):
        self.window = tk.Toplevel(parent)
        self.window.title(f"Tickets for Asset: {asset_id}")
        self.window.geometry("900x600")
        self.window.attributes('-topmost', True)
        
        ttk.Label(self.window, text=f"All tickets for asset: {asset_id}", 
                 font=('TkDefaultFont', 12, 'bold')).pack(pady=10)
        
        # Treeview
        tree_frame = ttk.Frame(self.window)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        self.tree = ttk.Treeview(tree_frame, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.tree.yview)
        
        scrollbar.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)
        
        self.tree['columns'] = ('ticket_id', 'type', 'created_date', 'created_by', 'status')
        self.tree.column('#0', width=0, stretch=False)
        
        for col in self.tree['columns']:
            self.tree.column(col, width=150)
            self.tree.heading(col, text=col.replace('_', ' ').title())
        
        # Load tickets
        self.load_tickets(asset_id)
        
        ttk.Button(self.window, text="Close", command=self.window.destroy).pack(pady=10)
    
    def load_tickets(self, asset_id):
        try:
            # Load ticket data
            with open(TIC_PATH, 'r') as f:
                ticket_data = json.load(f)
            
            for ticket in ticket_data.get('tickets', []):
                assets = ticket.get('assets', [])
                if not isinstance(assets, list):
                    assets = [ticket.get('asset', '')]
                
                if asset_id in assets:
                    self.tree.insert('', 'end', values=(
                        ticket.get('ticket_id', ''),
                        ticket.get('type', ''),
                        ticket.get('created_date', ''),
                        ticket.get('created_by', ''),
                        ticket.get('status', 'Open')
                    ))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load tickets: {str(e)}")

class ChangeLocationWindow:
    def __init__(self, parent, asset_ids, user, callback):
        self.asset_ids = asset_ids
        self.user = user
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("Change Location")
        self.window.geometry("500x300")
        self.window.attributes('-topmost', True)
        
        ttk.Label(self.window, text=f"Change location for {len(asset_ids)} asset(s)", 
                 font=('TkDefaultFont', 11, 'bold')).pack(pady=20)
        
        ttk.Label(self.window, text="Assets:", font=('TkDefaultFont', 10)).pack(pady=5)
        ttk.Label(self.window, text=', '.join(asset_ids), wraplength=400).pack(pady=5)
        
        ttk.Label(self.window, text="New Location:", font=('TkDefaultFont', 10)).pack(pady=10)
        self.location = ttk.Combobox(self.window, values=ASSET_SETTINGS.get('locations', []), 
                                     state='readonly', width=40)
        self.location.pack(pady=5)
        
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="Update", command=self.update_location).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def update_location(self):
        if not self.location.get():
            messagebox.showwarning("Warning", "Please select a location")
            return
        
        try:
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            for asset_id in self.asset_ids:
                for reg_item in registry_data:
                    if reg_item.get('identifier') == asset_id:
                        reg_item['location'] = self.location.get()
                        break
            
            save_json(REG_PATH, registry_data)
            
            self.callback()
            self.window.destroy()
            messagebox.showinfo("Success", f"Location updated for {len(self.asset_ids)} asset(s)")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update location: {str(e)}")

class ChangeStatusWindow:
    def __init__(self, parent, asset_ids, user, callback):
        self.asset_ids = asset_ids
        self.user = user
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("Change Status")
        self.window.geometry("500x300")
        self.window.attributes('-topmost', True)
        
        ttk.Label(self.window, text=f"Change status for {len(asset_ids)} asset(s)", 
                 font=('TkDefaultFont', 11, 'bold')).pack(pady=20)
        
        ttk.Label(self.window, text="Assets:", font=('TkDefaultFont', 10)).pack(pady=5)
        ttk.Label(self.window, text=', '.join(asset_ids), wraplength=400).pack(pady=5)
        
        ttk.Label(self.window, text="New Status:", font=('TkDefaultFont', 10)).pack(pady=10)
        self.status = ttk.Combobox(self.window, values=ASSET_SETTINGS.get('asset_statuses', []), 
                                   state='readonly', width=40)
        self.status.pack(pady=5)
        
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="Update", command=self.update_status).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def update_status(self):
        if not self.status.get():
            messagebox.showwarning("Warning", "Please select a status")
            return
        
        try:
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            for asset_id in self.asset_ids:
                for reg_item in registry_data:
                    if reg_item.get('identifier') == asset_id:
                        reg_item['status'] = self.status.get()
                        break
            
            save_json(REG_PATH, registry_data)
            
            self.callback()
            self.window.destroy()
            messagebox.showinfo("Success", f"Status updated for {len(self.asset_ids)} asset(s)")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update status: {str(e)}")

class AssetDetailsWindow:
    def __init__(self, parent, asset_id):
        self.window = tk.Toplevel(parent)
        self.window.title(f"Asset Details: {asset_id}")
        self.window.geometry("700x500")
        self.window.attributes('-topmost', True)
        
        # Load asset data
        try:
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            index_item = next((i for i in index_data if i.get('identifier') == asset_id), {})
            registry_item = next((r for r in registry_data if r.get('identifier') == asset_id), {})
            
            # Display info
            info_frame = ttk.Frame(self.window)
            info_frame.pack(fill='both', expand=True, padx=20, pady=20)
            
            ttk.Label(info_frame, text=f"Asset ID: {asset_id}", 
                     font=('TkDefaultFont', 14, 'bold')).pack(anchor='w', pady=10)
            
            ttk.Label(info_frame, text="Index Data (Read-Only):", 
                     font=('TkDefaultFont', 11, 'bold')).pack(anchor='w', pady=(10, 5))
            for key, value in index_item.items():
                if key != 'identifier':
                    ttk.Label(info_frame, text=f"{key}: {value}").pack(anchor='w', padx=20)
            
            ttk.Label(info_frame, text="Registry Data:", 
                     font=('TkDefaultFont', 11, 'bold')).pack(anchor='w', pady=(10, 5))
            for key, value in registry_item.items():
                if key != 'identifier':
                    ttk.Label(info_frame, text=f"{key}: {value}").pack(anchor='w', padx=20)
            
        except Exception as e:
            ttk.Label(self.window, text=f"Error loading asset data: {str(e)}").pack(pady=20)
        
        ttk.Button(self.window, text="Close", command=self.window.destroy).pack(pady=10)

class SettingsEditorWindow:
    """Live-editing settings window"""
    def __init__(self, parent, settings_path):
        self.settings_path = settings_path
        self.window = tk.Toplevel(parent)
        self.window.title("Asset Settings Editor")
        self.window.geometry("800x600")
        self.window.attributes('-topmost', True)
        
        # Load current settings
        try:
            with open(self.settings_path, 'r') as f:
                self.settings = json.load(f)
        except:
            self.settings = {}
        
        # Create notebook for different setting categories
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # General Settings Tab
        general_frame = ttk.Frame(notebook)
        notebook.add(general_frame, text="General")
        self.create_general_settings(general_frame)
        
        # Currencies Tab
        currencies_frame = ttk.Frame(notebook)
        notebook.add(currencies_frame, text="Currencies")
        self.create_currencies_settings(currencies_frame)
        
        # Delivery Addresses Tab
        addresses_frame = ttk.Frame(notebook)
        notebook.add(addresses_frame, text="Delivery Addresses")
        self.create_addresses_settings(addresses_frame)
        
        # Ignore Columns Tab
        ignore_frame = ttk.Frame(notebook)
        notebook.add(ignore_frame, text="Ignore Columns")
        self.create_ignore_columns_settings(ignore_frame)
        
        # Raw JSON Tab (for advanced editing)
        json_frame = ttk.Frame(notebook)
        notebook.add(json_frame, text="Raw JSON")
        self.create_json_editor(json_frame)
    
    def create_general_settings(self, parent):
        """Create general settings entries"""
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Ticket storage folder
        ttk.Label(scrollable_frame, text="Ticket Storage Folder:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ticket_folder_var = tk.StringVar(value=self.settings.get('ticket_storage_folder', ''))
        ttk.Entry(scrollable_frame, textvariable=ticket_folder_var, width=50).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ticket_folder_var.trace('w', lambda *args: self.save_setting('ticket_storage_folder', ticket_folder_var.get()))
        
        scrollable_frame.grid_columnconfigure(1, weight=1)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def create_currencies_settings(self, parent):
        """Create currencies list editor"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(frame, text="Currencies (Name | Symbol):").pack(anchor='w')
        
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='both', expand=True, pady=5)
        
        # Currency list
        currencies = self.settings.get('currencies', [])
        self.currency_listbox = tk.Listbox(list_frame)
        self.currency_listbox.pack(side='left', fill='both', expand=True)
        
        for curr in currencies:
            self.currency_listbox.insert('end', f"{curr.get('name', '')} | {curr.get('symbol', '')}")
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x')
        
        def add_currency():
            dialog = tk.Toplevel(self.window)
            dialog.title("Add Currency")
            dialog.attributes('-topmost', True)
            dialog.geometry("300x150")
            
            ttk.Label(dialog, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            name_var = tk.StringVar()
            ttk.Entry(dialog, textvariable=name_var).grid(row=0, column=1, padx=5, pady=5)
            
            ttk.Label(dialog, text="Symbol:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
            symbol_var = tk.StringVar()
            ttk.Entry(dialog, textvariable=symbol_var).grid(row=1, column=1, padx=5, pady=5)
            
            def save_currency():
                if name_var.get() and symbol_var.get():
                    currencies.append({'name': name_var.get(), 'symbol': symbol_var.get()})
                    self.currency_listbox.insert('end', f"{name_var.get()} | {symbol_var.get()}")
                    self.save_setting('currencies', currencies)
                    dialog.destroy()
            
            ttk.Button(dialog, text="Add", command=save_currency).grid(row=2, column=0, columnspan=2, pady=10)
        
        def remove_currency():
            selection = self.currency_listbox.curselection()
            if selection:
                idx = selection[0]
                del currencies[idx]
                self.currency_listbox.delete(idx)
                self.save_setting('currencies', currencies)
        
        ttk.Button(btn_frame, text="Add Currency", command=add_currency).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Remove Selected", command=remove_currency).pack(side='left', padx=2)
    
    def create_addresses_settings(self, parent):
        """Create delivery addresses list editor"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(frame, text="Delivery Addresses:").pack(anchor='w')
        
        # Address list
        addresses = self.settings.get('delivery_addresses', [])
        self.address_listbox = tk.Listbox(frame)
        self.address_listbox.pack(fill='both', expand=True, pady=5)
        
        for addr in addresses:
            self.address_listbox.insert('end', addr)
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x')
        
        def add_address():
            dialog = tk.Toplevel(self.window)
            dialog.title("Add Delivery Address")
            dialog.attributes('-topmost', True)
            dialog.geometry("400x150")
            
            ttk.Label(dialog, text="Address:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            address_var = tk.StringVar()
            ttk.Entry(dialog, textvariable=address_var, width=40).grid(row=0, column=1, padx=5, pady=5)
            
            def save_address():
                if address_var.get():
                    addresses.append(address_var.get())
                    self.address_listbox.insert('end', address_var.get())
                    self.save_setting('delivery_addresses', addresses)
                    dialog.destroy()
            
            ttk.Button(dialog, text="Add", command=save_address).grid(row=1, column=0, columnspan=2, pady=10)
        
        def remove_address():
            selection = self.address_listbox.curselection()
            if selection:
                idx = selection[0]
                del addresses[idx]
                self.address_listbox.delete(idx)
                self.save_setting('delivery_addresses', addresses)
        
        ttk.Button(btn_frame, text="Add Address", command=add_address).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Remove Selected", command=remove_address).pack(side='left', padx=2)
    
    def create_ignore_columns_settings(self, parent):
        """Create ignore columns list editor"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(frame, text="Columns to hide from Asset list:").pack(anchor='w')
        
        # Column list
        ignore_cols = self.settings.get('ignore_columns', [])
        self.ignore_listbox = tk.Listbox(frame)
        self.ignore_listbox.pack(fill='both', expand=True, pady=5)
        
        for col in ignore_cols:
            self.ignore_listbox.insert('end', col)
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x')
        
        def add_column():
            dialog = tk.Toplevel(self.window)
            dialog.title("Add Column to Ignore")
            dialog.attributes('-topmost', True)
            dialog.geometry("300x120")
            
            ttk.Label(dialog, text="Column Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            column_var = tk.StringVar()
            ttk.Entry(dialog, textvariable=column_var).grid(row=0, column=1, padx=5, pady=5)
            
            def save_column():
                if column_var.get():
                    ignore_cols.append(column_var.get())
                    self.ignore_listbox.insert('end', column_var.get())
                    self.save_setting('ignore_columns', ignore_cols)
                    dialog.destroy()
            
            ttk.Button(dialog, text="Add", command=save_column).grid(row=1, column=0, columnspan=2, pady=10)
        
        def remove_column():
            selection = self.ignore_listbox.curselection()
            if selection:
                idx = selection[0]
                del ignore_cols[idx]
                self.ignore_listbox.delete(idx)
                self.save_setting('ignore_columns', ignore_cols)
        
        ttk.Button(btn_frame, text="Add Column", command=add_column).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Remove Selected", command=remove_column).pack(side='left', padx=2)
    
    def create_json_editor(self, parent):
        """Create raw JSON editor"""
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(frame, text="Raw JSON (Advanced - Edit with caution):").pack(anchor='w')
        
        self.json_text = tk.Text(frame, wrap='word')
        self.json_text.pack(fill='both', expand=True, pady=5)
        
        # Load current JSON
        self.json_text.insert('1.0', json.dumps(self.settings, indent=2))
        
        def apply_json():
            try:
                new_settings = json.loads(self.json_text.get('1.0', 'end'))
                with open(self.settings_path, 'w') as f:
                    json.dump(new_settings, f, indent=2)
                self.settings = new_settings
                messagebox.showinfo("Success", "Settings saved")
            except Exception as e:
                messagebox.showerror("Error", f"Invalid JSON: {str(e)}")
        
        ttk.Button(frame, text="Apply JSON", command=apply_json).pack(pady=5)
    
    def save_setting(self, key, value):
        """Save a single setting immediately"""
        self.settings[key] = value
        try:
            with open(self.settings_path, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

class AddAssetWindow:
    """Window to add a new asset to index and registry"""
    def __init__(self, parent, callback):
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("Add New Asset")
        self.window.geometry("600x700")
        self.window.attributes('-topmost', True)
        
        # Load existing data to determine columns
        try:
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
        except:
            index_data = []
            registry_data = []
        
        # Get all possible columns
        index_cols = []
        if index_data and len(index_data) > 0:
            index_cols = list(index_data[0].keys())
        
        registry_cols = ['identifier', 'cal_date', 'cal_period', 'cal_by', 'next_cal_due', 
                        'check_period', 'last_checked', 'location', 'status']
        
        # Create scrollable form
        canvas = tk.Canvas(self.window)
        scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Store entry variables
        self.index_vars = {}
        self.registry_vars = {}
        
        row = 0
        
        # Index fields
        ttk.Label(scrollable_frame, text="Index Data (Read-Only Information):", 
                 font=('TkDefaultFont', 10, 'bold')).grid(row=row, column=0, columnspan=2, pady=10, sticky='w')
        row += 1
        
        for col in index_cols:
            ttk.Label(scrollable_frame, text=f"{col.replace('_', ' ').title()}:").grid(row=row, column=0, sticky='w', padx=5, pady=2)
            var = tk.StringVar()
            self.index_vars[col] = var
            ttk.Entry(scrollable_frame, textvariable=var, width=40).grid(row=row, column=1, sticky='ew', padx=5, pady=2)
            row += 1
        
        # Registry fields
        ttk.Label(scrollable_frame, text="Registry Data (Editable Information):", 
                 font=('TkDefaultFont', 10, 'bold')).grid(row=row, column=0, columnspan=2, pady=10, sticky='w')
        row += 1
        
        for col in registry_cols:
            if col != 'identifier':  # identifier is already in index
                ttk.Label(scrollable_frame, text=f"{col.replace('_', ' ').title()}:").grid(row=row, column=0, sticky='w', padx=5, pady=2)
                var = tk.StringVar()
                self.registry_vars[col] = var
                ttk.Entry(scrollable_frame, textvariable=var, width=40).grid(row=row, column=1, sticky='ew', padx=5, pady=2)
                row += 1
        
        scrollable_frame.grid_columnconfigure(1, weight=1)
        
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y", pady=10)
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Save Asset", command=self.save_asset).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def save_asset(self):
        """Save the new asset to both index and registry"""
        try:
            # Collect data
            index_item = {key: var.get() for key, var in self.index_vars.items()}
            registry_item = {key: var.get() for key, var in self.registry_vars.items()}
            
            # Ensure identifier is present
            if not index_item.get('identifier'):
                messagebox.showerror("Error", "Identifier is required")
                return
            
            # Add identifier to registry item
            registry_item['identifier'] = index_item['identifier']
            
            # Load existing data
            with open(INDEX_PATH, 'r') as f:
                index_data = json.load(f)
            
            with open(REG_PATH, 'r') as f:
                registry_data = json.load(f)
            
            # Check for duplicates
            if any(item.get('identifier') == index_item['identifier'] for item in index_data):
                messagebox.showerror("Error", "An asset with this identifier already exists")
                return
            
            # Add to data
            index_data.append(index_item)
            registry_data.append(registry_item)
            
            # Save with backups
            save_json(INDEX_PATH, index_data)
            save_json(REG_PATH, registry_data)
            
            messagebox.showinfo("Success", f"Asset {index_item['identifier']} added successfully")
            self.callback()
            self.window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add asset: {str(e)}")

class EditSupplierWindow:
    def __init__(self, parent, supplier_code, callback):
        self.supplier_code = supplier_code
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("Edit Supplier" if supplier_code else "Add Supplier")
        self.window.geometry("500x600")
        self.window.attributes('-topmost', True)
        
        self.entries = {}
        
        # Load existing data if editing
        if supplier_code:
            if SUPPLIER_LIST_PATH.exists():
                with open(SUPPLIER_LIST_PATH, 'r') as f:
                    suppliers_data = json.load(f)
                supplier = next((s for s in suppliers_data.get('suppliers', []) if s.get('supplier_code') == supplier_code), {})
            else:
                supplier = {}
        else:
            supplier = {}
        
        fields = ['supplier_code', 'name', 'contact', 'phone', 'email', 'address']
        row = 0

        for field in fields:
            ttk.Label(self.window, text=f"{field.replace('_', ' ').title()}:", font=('TkDefaultFont', 10)).grid(row=row, column=0, sticky='w', padx=10, pady=5)
            entry = ttk.Entry(self.window, width=40)
            entry.insert(0, str(supplier.get(field, '')))
            entry.grid(row=row, column=1, padx=10, pady=5)
            self.entries[field] = entry
            row += 1
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.window.destroy).pack(side='left', padx=5)
    
    def save(self):
        try:
            if not SUPPLIER_LIST_PATH.exists():
                suppliers_data = {"suppliers": []}
            else:
                with open(SUPPLIER_LIST_PATH, 'r') as f:
                    suppliers_data = json.load(f)
            
            suppliers = suppliers_data.get('suppliers', [])
            
            if self.supplier_code:
                # Edit existing supplier
                for i, sup in enumerate(suppliers):
                    if sup.get('supplier_code') == self.supplier_code:
                        for key, entry in self.entries.items():
                            suppliers[i][key] = entry.get()
                        break
            else:
                # Add new supplier
                new_supplier = {key: entry.get() for key, entry in self.entries.items()}
                suppliers.append(new_supplier)
            
            suppliers_data['suppliers'] = suppliers
            
            with open(SUPPLIER_LIST_PATH, 'w') as f:
                json.dump(suppliers_data, f, indent=4)
            
            messagebox.showinfo("Success", "Supplier saved successfully")
            self.callback()
            self.window.destroy()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save supplier: {str(e)}")

class SupplierWindow:
    def __init__(self, parent, ticket):
        self.window = tk.Toplevel(parent)
        self.window.title("Suppliers")
        self.window.geometry("800x600")
        self.window.attributes('-topmost', True)
        
        ttk.Label(self.window, text="Suppliers", 
                 font=('TkDefaultFont', 12, 'bold')).pack(pady=10)

        #Supplier List Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(padx=10, pady=5)
        ttk.Button(btn_frame, text="Add Supplier", command=self.add_supplier).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Edit Selected", command=self.edit_supplier).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Delete Selected", command=self.delete_supplier).pack(side='left', padx=5)

        # Supplier List Frame
        tree_frame = ttk.Frame(self.window)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient='vertical')
        scrollbar_x = ttk.Scrollbar(tree_frame, orient='horizontal')
        
        self.supplier_tree = ttk.Treeview(tree_frame, columns=('supplier_code', 'name', 'contact', 'phone', 'email', 'address'), 
                         show='headings', yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.config(command=self.supplier_tree.yview)
        scrollbar_x.config(command=self.supplier_tree.xview)
        
        for col in self.supplier_tree['columns']:
            self.supplier_tree.heading(col, text=col.replace('_', ' ').title())
            self.supplier_tree.column(col, width=120)
        
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.supplier_tree.pack(fill='both', expand=True)
        
        self.load_suppliers()

    def load_suppliers(self):
        """Load suppliers from settings"""
        self.supplier_tree.delete(*self.supplier_tree.get_children())
        suppliers = json.loads(SUPPLIER_LIST_PATH.read_text()).get('suppliers', [])
        for supplier in suppliers:
            self.supplier_tree.insert('', 'end', values=(
                supplier.get('supplier_code', ''),
                supplier.get('name', ''),
                supplier.get('contact', ''),
                supplier.get('phone', ''),
                supplier.get('email', ''),
                supplier.get('address', '')
            ))

    def add_supplier(self):
        EditSupplierWindow(self.window, None, self.load_suppliers)

    def edit_supplier(self):
        selected = self.supplier_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No supplier selected")
            return
        supplier_code = self.supplier_tree.item(selected[0])['values'][0]
        EditSupplierWindow(self.window, supplier_code, self.load_suppliers)
    
    def delete_supplier(self):
        selected = self.supplier_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No supplier selected")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this supplier?"):
            try:
                supplier_code = self.supplier_tree.item(selected[0])['values'][0]
                
                with open(SUPPLIER_LIST_PATH, 'r') as f:
                    suppliers_data = json.load(f)
                
                suppliers = suppliers_data.get('suppliers', [])
                suppliers_data['suppliers'] = [s for s in suppliers if s.get('supplier_code') != supplier_code]
                
                with open(SUPPLIER_LIST_PATH, 'w') as f:
                    json.dump(suppliers_data, f, indent=4)
                
                messagebox.showinfo("Success", "Supplier deleted successfully")
                self.load_suppliers()
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete supplier: {str(e)}")

root = tk.Tk()
app = AssetManagementApp(root)
root.mainloop()
