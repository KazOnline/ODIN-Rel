import json
import urllib.request
import urllib.error
import zipfile
import shutil
import os
from pathlib import Path
import sys

class UpdateChecker:
    def __init__(self):
        self.repo_owner = "KazOnline"
        self.repo_name = "ODIN-Rel"
        self.github_api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/latest"
        
        # Get current version from settings
        self.settings_path = Path(__file__).parent / "settings.json"
        self.current_version = self.get_current_version()
        self.latest_version = None
        self.download_zip_url = None
        
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
    
    def compare_versions(self, latest, current):
        """Compare version strings"""
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
    
    def check_for_updates(self):
        """Check GitHub for the latest release"""
        print("Checking for updates...")
        print(f"Current version: {self.current_version}")
        print(f"Contacting: {self.github_api_url}")
        
        try:
            # Request latest release info from GitHub API
            req = urllib.request.Request(self.github_api_url)
            req.add_header('User-Agent', 'ODIN-Updater')
            req.add_header('Accept', 'application/vnd.github+json')
            
            with urllib.request.urlopen(req, timeout=10) as response:
                print(f"Response code: {response.getcode()}")
                data = json.loads(response.read().decode())
            
            # Get version info
            self.latest_version = data.get('tag_name', '').lstrip('v')
            release_name = data.get('name', 'Unknown')
            
            print(f"Latest version: {self.latest_version}")
            
            # Try to get download URL from assets first, fallback to zipball
            assets = data.get('assets', [])
            print(f"Found {len(assets)} assets")
            if assets:
                # Look for a zip file in assets
                for asset in assets:
                    if asset.get('name', '').endswith('.zip'):
                        self.download_zip_url = asset.get('browser_download_url')
                        print(f"Using asset: {asset.get('name')}")
                        break
            
            # If no assets found, use zipball_url
            if not self.download_zip_url:
                self.download_zip_url = data.get('zipball_url')
                print(f"Using zipball URL")
            
            if not self.latest_version:
                raise Exception("No version information found in release")
            
            return True
                
        except urllib.error.HTTPError as e:
            if e.code == 404:
                print(f"ERROR: No releases found for this repository")
            else:
                print(f"ERROR: HTTP {e.code}")
            return False
        except urllib.error.URLError as e:
            print(f"ERROR: Connection error: {e.reason}")
            return False
        except Exception as e:
            print(f"ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def download_and_extract_update(self):
        """Download and extract the latest version"""
        if not self.download_zip_url:
            print("ERROR: No download URL available")
            return False
        
        print(f"\nDownloading version {self.latest_version}...")
        
        try:
            # Create updates directory
            update_dir = Path(__file__).parent / "updates"
            update_dir.mkdir(exist_ok=True)
            
            # Download zip file
            zip_path = update_dir / f"ODIN-{self.latest_version}.zip"
            print(f"Downloading to {zip_path}...")
            
            req = urllib.request.Request(self.download_zip_url)
            req.add_header('User-Agent', 'ODIN-Updater')
            
            with urllib.request.urlopen(req, timeout=60) as response:
                total_size = int(response.headers.get('content-length', 0))
                downloaded = 0
                chunk_size = 8192
                
                with open(zip_path, 'wb') as f:
                    while True:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            percent = (downloaded / total_size) * 100
                            print(f"\rProgress: {percent:.1f}%", end='')
            
            print("\n✓ Download complete!")
            
            # Extract zip
            extract_dir = update_dir / f"ODIN-{self.latest_version}"
            if extract_dir.exists():
                shutil.rmtree(extract_dir)
            
            print(f"Extracting to {extract_dir}...")
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            print("✓ Extraction complete!")
            
            # Copy files to current directory
            print("\nCopying files to application directory...")
            
            # Find the actual folder inside (GitHub adds repo name prefix)
            extracted_folders = [f for f in extract_dir.iterdir() if f.is_dir()]
            if extracted_folders:
                source_dir = extracted_folders[0]
            else:
                source_dir = extract_dir
            
            dest_dir = Path(__file__).parent
            
            # Copy all files except certain directories
            exclude_dirs = {'updates', 'temp', 'history', 'reports', '__pycache__'}
            exclude_files = {'database.json', 'log.json', 'stats.json', 'analytics.log', 'settings.json'}
            
            for item in source_dir.rglob('*'):
                if item.is_file():
                    rel_path = item.relative_to(source_dir)
                    
                    # Skip excluded directories
                    if any(part in exclude_dirs for part in rel_path.parts):
                        continue
                    
                    # Skip excluded files
                    if rel_path.name in exclude_files:
                        continue
                    
                    dest_file = dest_dir / rel_path
                    dest_file.parent.mkdir(parents=True, exist_ok=True)
                    
                    shutil.copy2(item, dest_file)
                    print(f"  Copied: {rel_path}")
            
            # Update version in settings
            if self.settings_path.exists():
                with open(self.settings_path, 'r') as f:
                    settings = json.load(f)
                settings['version'] = self.latest_version
                with open(self.settings_path, 'w') as f:
                    json.dump(settings, f, indent=4)
                print(f"\n✓ Updated version to {self.latest_version} in settings.json")
            
            print("\n✓ Update complete!")
            return True
            
        except Exception as e:
            print(f"\nERROR downloading update: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def run(self):
        """Run the update checker"""
        if not self.check_for_updates():
            print("\nFailed to check for updates")
            return False
        
        # Compare versions
        if self.compare_versions(self.latest_version, self.current_version):
            print(f"\n{'='*50}")
            print(f"UPDATE AVAILABLE!")
            print(f"Current: {self.current_version}")
            print(f"Latest:  {self.latest_version}")
            print(f"{'='*50}\n")
            
            # Force update
            if self.download_and_extract_update():
                print("\n✓ Update successfully installed!")
                print("Please restart the application.")
                return True
            else:
                print("\n✗ Update failed")
                return False
        else:
            print(f"\n✓ You are running the latest version ({self.current_version})")
            return True

if __name__ == "__main__":
    app = UpdateChecker()
    success = app.run()
    
    input("\nPress Enter to exit...")
    sys.exit(0 if success else 1)
