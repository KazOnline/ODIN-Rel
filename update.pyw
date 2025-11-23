import tkinter as tk
from tkinter import ttk, messagebox
import json
import urllib.request
import urllib.error
import zipfile
import shutil
import os
from pathlib import Path
import subprocess
import sys

class UpdateChecker:
    def __init__(self):
        self.repo_owner = "KazOnline"
        self.repo_name = "ODIN-Private"
        self.github_api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/latest"
        self.download_url = f"https://github.com/{self.repo_owner}/{self.repo_name}"
        
        # Get current version from settings
        self.settings_path = Path(__file__).parent / "settings.json"
        self.current_version = self.get_current_version()
        
        # Setup GUI
        self.root = tk.Tk()
        self.root.title("ODIN Update Checker")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # Try to load icon if it exists
        icon_path = Path(__file__).parent / "icons" / "icon.png"
        if icon_path.exists():
            try:
                icon = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(False, icon)
            except:
                pass
        
        self.setup_ui()
        
    def get_current_version(self):
        """Get current version from settings.json"""
        try:
            if self.settings_path.exists():
                with open(self.settings_path, 'r') as f:
                    settings = json.load(f)
                    return settings.get('version', '0.0.0')
        except Exception as e:
            print(f"Error reading version: {e}")
        return "0.0.0"
    
    def setup_ui(self):
        """Setup the user interface"""
        # Title
        title_label = ttk.Label(self.root, text="ODIN Update Checker", 
                               font=('TkDefaultFont', 16, 'bold'))
        title_label.pack(pady=20)
        
        # Current version
        version_frame = ttk.Frame(self.root)
        version_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Label(version_frame, text="Current Version:", 
                 font=('TkDefaultFont', 10, 'bold')).pack(side='left')
        self.current_ver_label = ttk.Label(version_frame, text=self.current_version,
                                           font=('TkDefaultFont', 10))
        self.current_ver_label.pack(side='left', padx=5)
        
        # Latest version
        latest_frame = ttk.Frame(self.root)
        latest_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Label(latest_frame, text="Latest Version:", 
                 font=('TkDefaultFont', 10, 'bold')).pack(side='left')
        self.latest_ver_label = ttk.Label(latest_frame, text="Checking...",
                                          font=('TkDefaultFont', 10))
        self.latest_ver_label.pack(side='left', padx=5)
        
        # Status text
        self.status_text = tk.Text(self.root, height=10, width=50, wrap='word', state='disabled')
        self.status_text.pack(pady=20, padx=20, fill='both', expand=True)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(pady=10, padx=20, fill='x')
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.check_btn = ttk.Button(button_frame, text="Check for Updates", 
                                    command=self.check_for_updates)
        self.check_btn.pack(side='left', padx=5)
        
        self.download_btn = ttk.Button(button_frame, text="Download Update", 
                                       command=self.download_update, state='disabled')
        self.download_btn.pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Close", 
                  command=self.root.quit).pack(side='left', padx=5)
        
        self.latest_version = None
        self.download_zip_url = None
        
    def log(self, message):
        """Add message to status text"""
        self.status_text.config(state='normal')
        self.status_text.insert('end', message + '\n')
        self.status_text.see('end')
        self.status_text.config(state='disabled')
        self.root.update()
        
    def check_for_updates(self):
        """Check GitHub for the latest release"""
        self.log("Checking for updates...")
        self.progress.start()
        self.check_btn.config(state='disabled')
        self.download_btn.config(state='disabled')
        
        try:
            # Request latest release info from GitHub API
            req = urllib.request.Request(self.github_api_url)
            req.add_header('User-Agent', 'ODIN-Updater')
            
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode())
                
            # Get version info
            self.latest_version = data.get('tag_name', '').lstrip('v')
            release_name = data.get('name', 'Unknown')
            release_notes = data.get('body', 'No release notes available.')
            
            # Get download URL for zip
            self.download_zip_url = data.get('zipball_url')
            
            self.latest_ver_label.config(text=self.latest_version)
            
            # Compare versions
            if self.compare_versions(self.latest_version, self.current_version):
                self.log(f"\n✓ New version available: {self.latest_version}")
                self.log(f"Release: {release_name}")
                self.log(f"\nRelease Notes:\n{release_notes[:200]}...")
                self.download_btn.config(state='normal')
                messagebox.showinfo("Update Available", 
                                  f"A new version ({self.latest_version}) is available!\n\n"
                                  f"Current: {self.current_version}\n"
                                  f"Latest: {self.latest_version}")
            else:
                self.log(f"\n✓ You are running the latest version ({self.current_version})")
                messagebox.showinfo("No Updates", 
                                  "You are already running the latest version!")
                
        except urllib.error.URLError as e:
            self.log(f"\n✗ Connection error: {e.reason}")
            messagebox.showerror("Connection Error", 
                               f"Could not connect to GitHub.\n{e.reason}")
        except Exception as e:
            self.log(f"\n✗ Error checking for updates: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        finally:
            self.progress.stop()
            self.check_btn.config(state='normal')
            
    def compare_versions(self, latest, current):
        """Compare version strings (simple comparison)"""
        try:
            latest_parts = [int(x) for x in latest.split('.')]
            current_parts = [int(x) for x in current.split('.')]
            
            # Pad shorter version with zeros
            while len(latest_parts) < len(current_parts):
                latest_parts.append(0)
            while len(current_parts) < len(latest_parts):
                current_parts.append(0)
                
            return latest_parts > current_parts
        except:
            return latest != current
            
    def download_update(self):
        """Download the latest version from GitHub"""
        if not self.download_zip_url:
            messagebox.showerror("Error", "No download URL available")
            return
            
        result = messagebox.askyesno("Confirm Download", 
                                     f"Download version {self.latest_version}?\n\n"
                                     "The update will be downloaded and extracted.\n"
                                     "You will need to manually replace files.")
        
        if not result:
            return
            
        self.log("\nDownloading update...")
        self.progress.start()
        self.download_btn.config(state='disabled')
        self.check_btn.config(state='disabled')
        
        try:
            # Create updates directory
            update_dir = Path(__file__).parent / "updates"
            update_dir.mkdir(exist_ok=True)
            
            # Download zip file
            zip_path = update_dir / f"ODIN-{self.latest_version}.zip"
            self.log(f"Downloading to {zip_path}...")
            
            req = urllib.request.Request(self.download_zip_url)
            req.add_header('User-Agent', 'ODIN-Updater')
            
            with urllib.request.urlopen(req, timeout=30) as response:
                with open(zip_path, 'wb') as f:
                    f.write(response.read())
                    
            self.log(f"✓ Download complete!")
            
            # Extract zip
            extract_dir = update_dir / f"ODIN-{self.latest_version}"
            self.log(f"Extracting to {extract_dir}...")
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                
            self.log("✓ Extraction complete!")
            self.log(f"\nUpdate files saved to:\n{extract_dir}")
            
            messagebox.showinfo("Download Complete", 
                              f"Update downloaded successfully!\n\n"
                              f"Files extracted to:\n{extract_dir}\n\n"
                              "Please manually copy the files to replace the old version.")
            
            # Open the folder
            if messagebox.askyesno("Open Folder", "Open the update folder?"):
                os.startfile(extract_dir)
                
        except Exception as e:
            self.log(f"\n✗ Error downloading update: {str(e)}")
            messagebox.showerror("Download Error", f"Failed to download update:\n{str(e)}")
        finally:
            self.progress.stop()
            self.check_btn.config(state='normal')
            
    def run(self):
        """Run the update checker"""
        self.root.mainloop()

if __name__ == "__main__":
    app = UpdateChecker()
    app.run()
