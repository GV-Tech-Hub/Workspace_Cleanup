# Standard library imports
import os
import logging
import json
import platform
import shutil
import time
import traceback
import datetime
import ctypes
import sys
import threading
import winreg
from pathlib import Path
from typing import Optional

# GUI imports
import tkinter as tk
from tkinter import ttk, messagebox

# Third-party imports
from PIL import Image, ImageDraw
import pystray

class FileOrganizer:
    def __init__(self, config_path: str = "config.json"):
        """Initialize the File Organizer with enhanced error handling"""
        try:
            # Set up logging first
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler('file_organizer.log'),
                    logging.StreamHandler()  # Also print to console
                ]
            )
            logging.info("Initializing File Organizer...")

            # Detect OS
            self.os_type = platform.system()
            logging.info(f"Detected OS: {self.os_type}")

            # Load configuration
            try:
                with open(config_path, 'r') as f:
                    self.config = json.load(f)
                logging.info("Configuration loaded successfully")
            except Exception as e:
                logging.error(f"Failed to load configuration: {e}")
                raise

            # Initialize statistics
            self.stats = {
                "files_moved": 0,
                "space_cleared": 0,
                "errors": 0
            }

            # Verify and get special folders with more detailed logging
            self.desktop_path = self.get_special_folder_path("Desktop")
            self.downloads_path = self.get_special_folder_path("Downloads")
            
            if not self.desktop_path:
                logging.error("Failed to get Desktop path")
            else:
                logging.info(f"Selected Desktop path: {self.desktop_path}")
                
            if not self.downloads_path:
                logging.error("Failed to get Downloads path")
            else:
                logging.info(f"Selected Downloads path: {self.downloads_path}")

            logging.info("Initialization complete")

        except Exception as e:
            logging.error(f"Initialization failed: {e}")
            logging.error(traceback.format_exc())
            raise

    def get_special_folder_path(self, folder_name: str) -> Optional[str]:
        """Get the path of special folders with multi-OS support"""
        try:
            if self.os_type == "Windows":
                # Define known folder GUIDs
                known_folders = {
                    'Desktop': '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}',
                    'Documents': '{FDD39AD0-238F-46AF-ADB4-6C85480369C7}',
                    'Downloads': '{374DE290-123F-4565-9164-39C4925E467B}',
                    'Music': '{4BD8D571-6D19-48D3-BE97-422220080E43}',
                    'Pictures': '{33E28130-4E1E-4676-835A-98395C3BC3BB}',
                    'Videos': '{18989B1D-99B5-455B-841C-AB7C74E4DDFC}'
                }

                if folder_name in known_folders:
                    try:
                        import winreg
                        sub_key = f"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders"
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                            path = winreg.QueryValueEx(key, folder_name)[0]
                            path = os.path.expandvars(path)
                            if os.path.exists(path):
                                return path
                    except Exception:
                        pass

                    # Try alternative method using shell32
                    try:
                        import ctypes
                        from ctypes import wintypes, windll
                        CSIDL_VALUES = {
                            'Desktop': 0,
                            'Documents': 5,
                            'Music': 13,
                            'Pictures': 39,
                            'Videos': 14,
                            'Downloads': 0x1A
                        }
                        if folder_name in CSIDL_VALUES:
                            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
                            windll.shell32.SHGetFolderPathW(None, CSIDL_VALUES[folder_name], None, 0, buf)
                            path = buf.value
                            if os.path.exists(path):
                                return path
                    except Exception:
                        pass

                # Fallback to environment variables
                env_vars = {
                    'Desktop': ['USERPROFILE', 'Desktop'],
                    'Documents': ['USERPROFILE', 'Documents'],
                    'Downloads': ['USERPROFILE', 'Downloads'],
                    'Music': ['USERPROFILE', 'Music'],
                    'Pictures': ['USERPROFILE', 'Pictures'],
                    'Videos': ['USERPROFILE', 'Videos']
                }

                if folder_name in env_vars:
                    base_path = os.environ.get(env_vars[folder_name][0], '')
                    path = os.path.join(base_path, env_vars[folder_name][1])
                    if os.path.exists(path):
                        return path

            elif self.os_type == "Darwin":  # macOS
                mac_paths = {
                    'Desktop': 'Desktop',
                    'Documents': 'Documents',
                    'Downloads': 'Downloads',
                    'Music': 'Music',
                    'Pictures': 'Pictures',
                    'Movies': 'Movies'  # macOS uses Movies instead of Videos
                }
                if folder_name in mac_paths:
                    path = os.path.join(os.path.expanduser('~'), mac_paths[folder_name])
                    if os.path.exists(path):
                        return path

            elif self.os_type == "Linux":
                # Use XDG user dirs
                try:
                    with open(os.path.expanduser('~/.config/user-dirs.dirs'), 'r') as f:
                        for line in f:
                            if line.startswith(f'XDG_{folder_name.upper()}_DIR'):
                                path = line.split('=')[1].strip('"').replace('$HOME', os.path.expanduser('~'))
                                if os.path.exists(path):
                                    return path
                except Exception:
                    pass

                # Fallback to standard XDG directories
                linux_paths = {
                    'Desktop': 'Desktop',
                    'Documents': 'Documents',
                    'Downloads': 'Downloads',
                    'Music': 'Music',
                    'Pictures': 'Pictures',
                    'Videos': 'Videos'
                }
                if folder_name in linux_paths:
                    path = os.path.join(os.path.expanduser('~'), linux_paths[folder_name])
                    if os.path.exists(path):
                        return path

            logging.info(f"Using {self.os_type} {folder_name} path: {path}")
            return path
        except Exception as e:
            logging.error(f"Error getting {folder_name} path: {e}")
            logging.error(traceback.format_exc())
            return None

    def get_file_size(self, file_path: str) -> int:
        """Get file size in bytes"""
        try:
            return os.path.getsize(file_path)
        except Exception:
            return 0

    def is_system_file(self, file_path: str) -> bool:
        """Enhanced system file checker"""
        try:
            file_name = os.path.basename(file_path)
            file_lower = file_name.lower()
            
            # Expanded system files list
            system_files = [
                "desktop.ini",
                "recycle bin",
                "trash",
                "$recycle.bin",
                ".ds_store",  # macOS system file
                "thumbs.db"    # Windows thumbnail cache
            ]
            
            # Check if it's a system file
            if file_lower in system_files:
                return True
            
            # Check if it's the Recycle Bin folder
            if os.path.isdir(file_path) and ("$recycle.bin" in file_lower or "recycle bin" in file_lower):
                return True
            
            return False
        except Exception as e:
            logging.warning(f"Error checking system file {file_path}: {e}")
            return False

    def update_archive_timestamp(self, archive_path: str) -> str:
        """Update existing archive folder name with current timestamp"""
        try:
            if os.path.exists(archive_path):
                current_time = datetime.datetime.now()
                new_timestamp = current_time.strftime("%b-%d-%Y_%I-%M%p")
                parent_dir = os.path.dirname(archive_path)
                new_path = os.path.join(parent_dir, f"Archive_{new_timestamp}")
                
                # Rename the existing archive folder
                try:
                    os.rename(archive_path, new_path)
                    logging.info(f"Updated archive timestamp: {os.path.basename(new_path)}")
                    return new_path
                except Exception as e:
                    logging.error(f"Failed to rename archive folder: {e}")
                    return archive_path
        except Exception as e:
            logging.error(f"Error updating archive timestamp: {e}")
            return archive_path

    def set_folder_color(self, folder_path: str, category: str) -> None:
        """Set folder color and icon using Desktop.ini"""
        try:
            if self.os_type != "Windows":
                return

            ini_path = os.path.join(folder_path, "Desktop.ini")
            with open(ini_path, 'w') as f:
                f.write("[.ShellClassInfo]\n")
                
                # Category-specific icons
                if category == "Shortcuts":
                    f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,29\n")  # Star/favorite icon
                    f.write("IconFile=%SystemRoot%\\system32\\SHELL32.dll\n")
                    f.write("IconIndex=29\n")
                    f.write("BackgroundColor=255 215 0\n")  # Gold
                # ... (rest of your categories)

            # Set system and hidden attributes
            ctypes.windll.kernel32.SetFileAttributesW(ini_path, 0x2 | 0x4)
            ctypes.windll.kernel32.SetFileAttributesW(folder_path, 0x1)

        except Exception as e:
            logging.warning(f"Failed to set folder icon and color: {e}")

    def get_onedrive_archives(self) -> list:
        """Get all OneDrive archive paths"""
        archive_paths = []
        try:
            # Check for OneDrive Business
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                              r"Software\Microsoft\OneDrive\Accounts\Business1") as key:
                onedrive_path = winreg.QueryValueEx(key, "UserFolder")[0]
                archive_path = os.path.join(onedrive_path, "Archive - Pre-2024")
                if os.path.exists(archive_path):
                    archive_paths.append(archive_path)
        except Exception:
            pass

        try:
            # Check for OneDrive Personal
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                              r"Software\Microsoft\OneDrive\Accounts\Personal") as key:
                onedrive_path = winreg.QueryValueEx(key, "UserFolder")[0]
                archive_path = os.path.join(onedrive_path, "Archive - Pre-2024")
                if os.path.exists(archive_path):
                    archive_paths.append(archive_path)
        except Exception:
            pass

        return archive_paths

    def update_all_archives(self) -> None:
        """Find and update all archive folders recursively"""
        try:
            # Get all possible archive locations
            archive_locations = []
            
            # Add OneDrive archives
            archive_locations.extend(self.get_onedrive_archives())
            
            # Add local archives
            if self.desktop_path:
                archive_locations.append(os.path.join(self.desktop_path, "Archive"))
            if self.downloads_path:
                archive_locations.append(os.path.join(self.downloads_path, "Archive"))

            for archive_path in archive_locations:
                if os.path.exists(archive_path):
                    logging.info(f"Processing archive: {archive_path}")
                    
                    # Process all dated folders
                    for root, dirs, _ in os.walk(archive_path):
                        for dir_name in dirs:
                            if "_" in dir_name:  # Likely a dated folder
                                dated_path = os.path.join(root, dir_name)
                                logging.info(f"Updating dated folder: {dated_path}")
                                
                                # Create/update category folders
                                for category in self.categories.keys():
                                    category_path = os.path.join(dated_path, category)
                                    os.makedirs(category_path, exist_ok=True)
                                    self.set_folder_color(category_path, category)
                                
                                # Move misplaced files
                                for item in os.listdir(dated_path):
                                    item_path = os.path.join(dated_path, item)
                                    if os.path.isfile(item_path):
                                        self.move_to_category(item_path, dated_path)

        except Exception as e:
            logging.error(f"Failed to update archives: {e}")
            logging.error(traceback.format_exc())

    def move_to_category(self, file_path: str, dated_folder: str) -> None:
        """Move a file to its correct category folder"""
        try:
            file_name = os.path.basename(file_path)
            _, ext = os.path.splitext(file_name.lower())
            
            # Determine category
            category = 'Others'
            
            # Check for shortcuts first
            if ext in ['.lnk', '.url', '.desktop']:
                category = 'Shortcuts'
            else:
                # Check other categories
                for cat, info in self.categories.items():
                    if ext in info['extensions']:
                        category = cat
                        break
            
            target_dir = os.path.join(dated_folder, category)
            target_path = os.path.join(target_dir, file_name)
            
            if not os.path.exists(target_path):
                shutil.move(file_path, target_path)
                logging.info(f"Moved {file_name} to {category}")
            
        except Exception as e:
            logging.error(f"Failed to move file {file_path}: {e}")

    def organize_folder(self, folder_path: str) -> None:
        """Organize files with enhanced archive handling and category sorting"""
        try:
            if not folder_path or not os.path.exists(folder_path):
                logging.error(f"Invalid or non-existent folder path: {folder_path}")
                return

            # Get all items in the folder
            items = os.listdir(folder_path)
            
            # Handle shortcuts first
            shortcuts = [item for item in items if item.lower().endswith(('.lnk', '.url', '.desktop'))]
            if shortcuts:
                timestamp = datetime.datetime.now().strftime("%b-%d-%Y_%I-%M%p")
                archive_path = os.path.join(folder_path, "Archive", timestamp)
                shortcuts_path = os.path.join(archive_path, "Shortcuts")
                os.makedirs(shortcuts_path, exist_ok=True)
                self.set_folder_color(shortcuts_path, "Shortcuts")
                
                for shortcut in shortcuts:
                    try:
                        source_path = os.path.join(folder_path, shortcut)
                        target_path = os.path.join(shortcuts_path, shortcut)
                        shutil.move(source_path, target_path)
                        logging.info(f"Moved shortcut: {shortcut}")
                    except Exception as e:
                        logging.error(f"Failed to move shortcut {shortcut}: {e}")

            # Continue with regular organization
            # ... (rest of the organize_folder code remains the same)

        except Exception as e:
            logging.error(f"Failed to organize folder {folder_path}: {e}")
            logging.error(traceback.format_exc())

    def handle_shortcuts_first(self, folder_path: str) -> None:
        """Aggressively handle shortcuts before anything else"""
        try:
            # Force create Shortcuts folder in current archive
            timestamp = datetime.datetime.now().strftime("%b-%d-%Y_%I-%M%p")
            archive_path = os.path.join(folder_path, "Archive", timestamp)
            shortcuts_path = os.path.join(archive_path, "Shortcuts")
            os.makedirs(shortcuts_path, exist_ok=True)
            
            # Set shortcuts folder icon and color
            ini_path = os.path.join(shortcuts_path, "Desktop.ini")
            with open(ini_path, 'w', encoding='utf-8') as f:
                f.write("[.ShellClassInfo]\n")
                f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,29\n")
                f.write("IconFile=%SystemRoot%\\system32\\SHELL32.dll\n")
                f.write("IconIndex=29\n")
                f.write("BackgroundColor=255 215 0\n")  # Gold
            
            # Set folder attributes
            ctypes.windll.kernel32.SetFileAttributesW(ini_path, 0x2 | 0x4)  # Hidden | System
            ctypes.windll.kernel32.SetFileAttributesW(shortcuts_path, 0x1)  # Read-only
            
            # Find and move all shortcuts
            for item in os.listdir(folder_path):
                if item.lower().endswith(('.lnk', '.url', '.desktop')):
                    try:
                        source = os.path.join(folder_path, item)
                        target = os.path.join(shortcuts_path, item)
                        if os.path.exists(source):
                            shutil.move(source, target)
                            print(f"Moved shortcut: {item}")  # Direct console output
                            logging.info(f"Moved shortcut: {item}")
                    except Exception as e:
                        print(f"Failed to move shortcut {item}: {e}")  # Direct console output
                        logging.error(f"Failed to move shortcut {item}: {e}")
            
        except Exception as e:
            print(f"Error in handle_shortcuts_first: {e}")  # Direct console output
            logging.error(f"Error in handle_shortcuts_first: {e}")
            logging.error(traceback.format_exc())

    def force_update_onedrive_archive(self) -> None:
        """Force update OneDrive archives"""
        try:
            # Try both possible OneDrive paths
            onedrive_paths = [
                os.path.join(os.path.expanduser('~'), 'OneDrive', 'Archive - Pre-2024'),
                os.path.join(os.path.expanduser('~'), 'OneDrive - Business', 'Archive - Pre-2024')
            ]
            
            for onedrive_path in onedrive_paths:
                if os.path.exists(onedrive_path):
                    print(f"Updating OneDrive archive: {onedrive_path}")  # Direct console output
                    
                    # Process each dated folder
                    for item in os.listdir(onedrive_path):
                        dated_path = os.path.join(onedrive_path, item)
                        if os.path.isdir(dated_path) and "_" in item:
                            print(f"Processing: {item}")  # Direct console output
                            
                            # Update/create each category folder
                            for category, info in self.categories.items():
                                category_path = os.path.join(dated_path, category)
                                os.makedirs(category_path, exist_ok=True)
                                
                                # Set folder icon and color
                                ini_path = os.path.join(category_path, "Desktop.ini")
                                with open(ini_path, 'w', encoding='utf-8') as f:
                                    f.write("[.ShellClassInfo]\n")
                                    if category == "Audio":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,137\n")
                                        f.write("BackgroundColor=144 238 144\n")  # Light green
                                    elif category == "Documents":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,21\n")
                                        f.write("BackgroundColor=255 232 186\n")  # Light yellow
                                    elif category == "Images":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,140\n")
                                        f.write("BackgroundColor=173 216 230\n")  # Light blue
                                    # ... (add other categories as needed)
                                
                                # Set folder attributes
                                ctypes.windll.kernel32.SetFileAttributesW(ini_path, 0x2 | 0x4)
                                ctypes.windll.kernel32.SetFileAttributesW(category_path, 0x1)
                                
        except Exception as e:
            print(f"Error updating OneDrive archive: {e}")  # Direct console output
            logging.error(f"Error updating OneDrive archive: {e}")
            logging.error(traceback.format_exc())

    def move_shortcuts(self, folder_path: str) -> None:
        """Move all shortcuts to a dedicated folder."""
        try:
            timestamp = datetime.datetime.now().strftime("%b-%d-%Y_%I-%M%p")
            archive_path = os.path.join(folder_path, "Archive", timestamp)
            shortcuts_path = os.path.join(archive_path, "Shortcuts")
            os.makedirs(shortcuts_path, exist_ok=True)
            
            for item in os.listdir(folder_path):
                if item.lower().endswith(('.lnk', '.url', '.desktop')):
                    source = os.path.join(folder_path, item)
                    target = os.path.join(shortcuts_path, item)
                    shutil.move(source, target)
                    logging.info(f"Moved shortcut: {item}")
        except Exception as e:
            logging.error(f"Failed to move shortcuts: {e}")

    def update_existing_archives(self, root_path: str) -> None:
        """Update all existing archive folders."""
        try:
            archive_path = os.path.join(root_path, "Archive")
            if os.path.exists(archive_path):
                for dated_folder in os.listdir(archive_path):
                    dated_path = os.path.join(archive_path, dated_folder)
                    if os.path.isdir(dated_path):
                        for category in self.categories.keys():
                            category_path = os.path.join(dated_path, category)
                            os.makedirs(category_path, exist_ok=True)
                            self.set_folder_icon_and_color(category_path, category)
        except Exception as e:
            logging.error(f"Failed to update existing archives: {e}")

    def set_folder_icon_and_color(self, folder_path: str, category: str) -> None:
        """Set folder icon and color."""
        try:
            ini_path = os.path.join(folder_path, "Desktop.ini")
            with open(ini_path, 'w', encoding='utf-8') as f:
                f.write("[.ShellClassInfo]\n")
                if category == "Audio":
                    f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,137\n")
                    f.write("BackgroundColor=144 238 144\n")  # Light green
                elif category == "Documents":
                    f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,21\n")
                    f.write("BackgroundColor=255 232 186\n")  # Light yellow
                elif category == "Images":
                    f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,140\n")
                    f.write("BackgroundColor=173 216 230\n")  # Light blue
                elif category == "Shortcuts":
                    f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,29\n")
                    f.write("BackgroundColor=255 215 0\n")  # Gold
                # Add other categories as needed

            ctypes.windll.kernel32.SetFileAttributesW(ini_path, 0x2 | 0x4)
            ctypes.windll.kernel32.SetFileAttributesW(folder_path, 0x1)
        except Exception as e:
            logging.error(f"Failed to set icon and color for {folder_path}: {e}")

    def fix_all_archives(self) -> None:
        """Fix all existing archives, including OneDrive archives"""
        try:
            # List of all possible archive locations
            archive_locations = []
            
            # Add OneDrive locations
            onedrive_paths = [
                os.path.join(os.path.expanduser('~'), 'OneDrive', 'Archive - Pre-2024'),
                os.path.join(os.path.expanduser('~'), 'OneDrive - Business', 'Archive - Pre-2024')
            ]
            archive_locations.extend(onedrive_paths)
            
            # Add local archive locations
            if self.desktop_path:
                archive_locations.append(os.path.join(self.desktop_path, "Archive"))
            if self.downloads_path:
                archive_locations.append(os.path.join(self.downloads_path, "Archive"))

            # Process each archive location
            for archive_path in archive_locations:
                if os.path.exists(archive_path):
                    print(f"Fixing archive: {archive_path}")
                    
                    # Walk through all subfolders
                    for root, dirs, files in os.walk(archive_path):
                        # Check if this is a dated folder (contains underscore)
                        if os.path.basename(root).count('_') > 0:
                            print(f"Processing dated folder: {root}")
                            
                            # Ensure all category folders exist
                            for category in self.categories.keys():
                                category_path = os.path.join(root, category)
                                os.makedirs(category_path, exist_ok=True)
                                
                                # Set correct icon and color
                                ini_path = os.path.join(category_path, "Desktop.ini")
                                with open(ini_path, 'w', encoding='utf-8') as f:
                                    f.write("[.ShellClassInfo]\n")
                                    
                                    # Set category-specific icons and colors
                                    if category == "Audio":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,137\n")
                                        f.write("BackgroundColor=144 238 144\n")  # Light green
                                    elif category == "Documents":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,21\n")
                                        f.write("BackgroundColor=255 232 186\n")  # Light yellow
                                    elif category == "Images":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,140\n")
                                        f.write("BackgroundColor=173 216 230\n")  # Light blue
                                    elif category == "Video":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,136\n")
                                        f.write("BackgroundColor=255 182 193\n")  # Light pink
                                    elif category == "Shortcuts":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,29\n")
                                        f.write("BackgroundColor=255 215 0\n")  # Gold
                                    elif category == "Code":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,70\n")
                                        f.write("BackgroundColor=221 160 221\n")  # Purple
                                    elif category == "Executables":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,8\n")
                                        f.write("BackgroundColor=255 160 122\n")  # Coral
                                    elif category == "ZIP_Files":
                                        f.write("IconResource=%SystemRoot%\\system32\\zipfldr.dll,0\n")
                                        f.write("BackgroundColor=210 180 140\n")  # Tan
                                    elif category == "RAR_Files":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,165\n")
                                        f.write("BackgroundColor=210 180 140\n")  # Tan
                                    elif category == "Other_Archives":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,165\n")
                                        f.write("BackgroundColor=210 180 140\n")  # Tan
                                    else:  # Others
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,234\n")
                                        f.write("BackgroundColor=211 211 211\n")  # Light gray

                                # Set proper attributes
                                ctypes.windll.kernel32.SetFileAttributesW(ini_path, 0x2 | 0x4)  # Hidden | System
                                ctypes.windll.kernel32.SetFileAttributesW(category_path, 0x1)  # Read-only
                                
                                # Move any misplaced files into correct categories
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    if os.path.isfile(file_path):
                                        _, ext = os.path.splitext(file.lower())
                                        target_category = self.get_category_for_extension(ext)
                                        if target_category:
                                            target_path = os.path.join(root, target_category, file)
                                            if not os.path.exists(target_path):
                                                shutil.move(file_path, target_path)
                                                print(f"Moved {file} to {target_category}")

        except Exception as e:
            print(f"Error fixing archives: {e}")
            logging.error(f"Error fixing archives: {e}")
            logging.error(traceback.format_exc())

    def get_category_for_extension(self, ext: str) -> str:
        """Get the appropriate category for a file extension"""
        if ext in ['.lnk', '.url', '.desktop']:
            return 'Shortcuts'
        
        for category, info in self.categories.items():
            if ext in info['extensions']:
                return category
        return 'Others'

    def organize(self):
        """Main organization method"""
        try:
            print("Starting organization process...")
            logging.info("Starting organization process...")
            
            # Fix OneDrive archives first
            print("Fixing OneDrive archives...")
            self.fix_onedrive_archives()
            
            # Handle current organization
            if self.desktop_path:
                print("Organizing Desktop...")
                self.organize_folder(self.desktop_path)
            
            if self.downloads_path:
                print("Organizing Downloads...")
                self.organize_folder(self.downloads_path)
            
            # Fix OneDrive archives again
            print("Final OneDrive archive update...")
            self.fix_onedrive_archives()
            
            print("Organization complete!")
            self.print_statistics()
            
        except Exception as e:
            print(f"Organization failed: {e}")
            logging.error(f"Organization failed: {e}")
            logging.error(traceback.format_exc())

    def print_statistics(self):
        """Print organization statistics"""
        logging.info("\n=== Organization Statistics ===")
        logging.info(f"Files moved: {self.stats['files_moved']}")
        logging.info(f"Space cleared: {self.stats['space_cleared'] / (1024*1024):.2f} MB")
        logging.info(f"Errors encountered: {self.stats['errors']}")
        logging.info("===========================")

    def fix_onedrive_archives(self) -> None:
        """Fix OneDrive archives specifically"""
        try:
            # OneDrive paths with full structure
            onedrive_paths = [
                os.path.join(os.path.expanduser('~'), 'OneDrive', 'Archive - Pre-2024', 'Desktop', 'Archive'),
                os.path.join(os.path.expanduser('~'), 'OneDrive', 'Archive - Pre-2024', 'Downloads', 'Archive'),
                os.path.join(os.path.expanduser('~'), 'OneDrive - Business', 'Archive - Pre-2024', 'Desktop', 'Archive'),
                os.path.join(os.path.expanduser('~'), 'OneDrive - Business', 'Archive - Pre-2024', 'Downloads', 'Archive')
            ]

            for archive_path in onedrive_paths:
                if os.path.exists(archive_path):
                    print(f"Processing OneDrive archive: {archive_path}")
                    
                    # Process each dated folder
                    for dated_folder in os.listdir(archive_path):
                        dated_path = os.path.join(archive_path, dated_folder)
                        if os.path.isdir(dated_path) and '_' in dated_folder:
                            print(f"Fixing dated folder: {dated_folder}")
                            
                            # Ensure all category folders exist with correct icons
                            for category in self.categories.keys():
                                category_path = os.path.join(dated_path, category)
                                os.makedirs(category_path, exist_ok=True)
                                
                                # Set icon and color
                                ini_path = os.path.join(category_path, "Desktop.ini")
                                with open(ini_path, 'w', encoding='utf-8') as f:
                                    f.write("[.ShellClassInfo]\n")
                                    
                                    if category == "Audio":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,137\n")
                                        f.write("BackgroundColor=144 238 144\n")  # Light green
                                    elif category == "Documents":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,21\n")
                                        f.write("BackgroundColor=255 232 186\n")  # Light yellow
                                    elif category == "Images":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,140\n")
                                        f.write("BackgroundColor=173 216 230\n")  # Light blue
                                    elif category == "Video":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,136\n")
                                        f.write("BackgroundColor=255 182 193\n")  # Light pink
                                    elif category == "Shortcuts":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,29\n")
                                        f.write("BackgroundColor=255 215 0\n")  # Gold
                                    elif category == "Code":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,70\n")
                                        f.write("BackgroundColor=221 160 221\n")  # Purple
                                    elif category == "Executables":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,8\n")
                                        f.write("BackgroundColor=255 160 122\n")  # Coral
                                    elif category == "ZIP_Files":
                                        f.write("IconResource=%SystemRoot%\\system32\\zipfldr.dll,0\n")
                                        f.write("BackgroundColor=210 180 140\n")  # Tan
                                    elif category == "RAR_Files":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,165\n")
                                        f.write("BackgroundColor=210 180 140\n")  # Tan
                                    elif category == "Other_Archives":
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,165\n")
                                        f.write("BackgroundColor=210 180 140\n")  # Tan
                                    else:  # Others
                                        f.write("IconResource=%SystemRoot%\\system32\\SHELL32.dll,234\n")
                                        f.write("BackgroundColor=211 211 211\n")  # Light gray

                                # Set attributes
                                ctypes.windll.kernel32.SetFileAttributesW(ini_path, 0x2 | 0x4)  # Hidden | System
                                ctypes.windll.kernel32.SetFileAttributesW(category_path, 0x1)  # Read-only

                                # Move any misplaced files
                                for item in os.listdir(dated_path):
                                    if os.path.isfile(os.path.join(dated_path, item)):
                                        _, ext = os.path.splitext(item.lower())
                                        target_category = self.get_category_for_extension(ext)
                                        if target_category:
                                            source = os.path.join(dated_path, item)
                                            target = os.path.join(dated_path, target_category, item)
                                            if not os.path.exists(target):
                                                shutil.move(source, target)
                                                print(f"Moved {item} to {target_category}")

        except Exception as e:
            print(f"Error fixing OneDrive archives: {e}")
            logging.error(f"Error fixing OneDrive archives: {e}")
            logging.error(traceback.format_exc())

class FileOrganizerGUI:
    def __init__(self):
        """Initialize the GUI with custom styling"""
        try:
            # Initialize main window
            self.window = tk.Tk()
            self.window.title("DDA")
            
            # Configure window
            self.window.geometry("44x44+{}+{}".format(
                self.window.winfo_screenwidth() - 60,
                20
            ))
            
            # Window properties
            self.window.overrideredirect(True)
            self.window.attributes('-topmost', True)
            self.window.attributes('-alpha', 0.95)
            self.window.configure(bg='#006400')
            
            # Create frame for padding and background
            frame = tk.Frame(
                self.window,
                bg='#006400',
                padx=2,
                pady=2
            )
            frame.pack(fill='both', expand=True)
            
            # Create custom round cleanup button
            self.button = tk.Button(
                frame,
                text="🧹",
                font=('Segoe UI Emoji', 12),
                width=2,
                height=1,
                bg='#FF4444',
                fg='#FFD700',
                relief='raised',
                bd=1,
                highlightthickness=1,
                highlightbackground='#FFD700',
                cursor='hand2',
                command=self.show_organizer_dialog
            )
            self.button.pack(padx=2, pady=2)
            
            # Bind button events
            self.button.bind('<Button-1>', self.handle_click)
            self.button.bind('<B1-Motion>', self.on_move)
            self.button.bind('<Button-3>', self.show_menu)
            self.button.bind('<Enter>', lambda e: self.button.config(bg='#FF6666'))
            self.button.bind('<Leave>', lambda e: self.button.config(bg='#FF4444'))
            
            # Create right-click menu
            self.menu = tk.Menu(
                self.window,
                tearoff=0,
                bg='#006400',
                fg='#FFD700',
                activebackground='#008000',
                activeforeground='#FFFF00'
            )
            self.menu.add_command(label="Exit", command=self.exit_app)
            
            # Initialize drag variables
            self.drag_start_x = 0
            self.drag_start_y = 0
            self.dragging = False
            
            # Initialize tray icon
            self.tray_icon = None
            self.setup_tray()
            
            # Keep window open
            self.window.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
            
        except Exception as e:
            logging.error(f"GUI initialization error: {e}")
            logging.error(traceback.format_exc())
            raise

    def setup_tray(self):
        """Setup system tray icon"""
        try:
            # Create a simple icon (16x16 pixels)
            icon = Image.new('RGBA', (16, 16), color=(0, 100, 0, 0))
            d = ImageDraw.Draw(icon)
            d.rectangle([0, 0, 15, 15], outline='yellow')
            d.text((4, 2), "🧹", fill='yellow', font=None)

            def show_window(icon, item):
                self.window.deiconify()

            def exit_app(icon, item):
                icon.stop()
                self.window.quit()

            # Create system tray icon
            self.tray_icon = pystray.Icon(
                "DDA",
                icon,
                "Desktop Downloads Archiver",
                menu=pystray.Menu(
                    pystray.MenuItem("Show", show_window),
                    pystray.MenuItem("Exit", exit_app)
                )
            )
            
            # Run tray icon in separate thread
            threading.Thread(target=self.run_tray, daemon=True).start()
            
        except Exception as e:
            logging.error(f"Tray setup error: {e}")
            # Continue without tray icon
            pass

    def run_tray(self):
        """Run tray icon in separate thread"""
        try:
            self.tray_icon.run()
        except Exception as e:
            logging.error(f"Tray icon error: {e}")

    def minimize_to_tray(self):
        """Minimize window instead of closing"""
        self.window.withdraw()

    def exit_app(self):
        """Properly close the application"""
        try:
            if self.tray_icon:
                self.tray_icon.stop()
            self.window.quit()
        except Exception as e:
            logging.error(f"Exit error: {e}")
            sys.exit(1)

    def handle_click(self, event):
        """Handle button click - either start drag or show dialog"""
        if event.num == 1:  # Left click
            # Store initial position for potential drag
            self.drag_start_x = event.x_root - self.window.winfo_x()
            self.drag_start_y = event.y_root - self.window.winfo_y()
            self.dragging = False
            # Show dialog only on button release
            self.button.bind('<ButtonRelease-1>', self.check_click_or_drag)

    def check_click_or_drag(self, event):
        """Determine if this was a click or drag"""
        self.button.unbind('<ButtonRelease-1>')
        if not self.dragging:
            self.show_organizer_dialog()

    def show_organizer_dialog(self):
        """Show the folder selection dialog"""
        try:
            dialog = tk.Toplevel(self.window)
            dialog.title("Organize Folders")
            dialog.geometry("400x500")  # Made taller for better spacing
            dialog.configure(bg='#006400')  # Dark green background
            
            # Make dialog stay on top
            dialog.transient(self.window)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
            dialog.geometry(f'{width}x{height}+{x}+{y}')
            
            # Header label
            tk.Label(
                dialog,
                text="Select Folders to Organize",
                font=('Arial', 16, 'bold'),
                bg='#006400',
                fg='#FFD700'  # Yellow text
            ).pack(pady=20)
            
            # Folders frame
            folders_frame = tk.Frame(dialog, bg='#006400', padx=30)
            folders_frame.pack(fill='x')
            
            folders = {
                'Desktop': tk.BooleanVar(value=True),
                'Downloads': tk.BooleanVar(value=True),
                'Documents': tk.BooleanVar(value=False),
                'Music': tk.BooleanVar(value=False),
                'Pictures': tk.BooleanVar(value=False),
                'Videos': tk.BooleanVar(value=False),
                'Workspaces': tk.BooleanVar(value=False)
            }
            
            for folder, var in folders.items():
                cb = tk.Checkbutton(
                    folders_frame,
                    text=folder,
                    variable=var,
                    bg='#006400',
                    fg='#FFD700',  # Yellow text
                    selectcolor='#004d00',
                    activebackground='#006400',
                    activeforeground='#FFFF00',
                    font=('Arial', 12),
                    width=20,
                    anchor='w'
                )
                cb.pack(anchor='w', pady=8)
            
            # Button frame for better positioning
            button_frame = tk.Frame(dialog, bg='#006400')
            button_frame.pack(side='bottom', pady=30)
            
            # Start button with updated styling
            start_button = tk.Button(
                button_frame,
                text="START CLEANUP",
                command=lambda: self.start_cleanup(dialog, folders),
                bg='#FFD700',  # Yellow background
                fg='#006400',  # Dark green text
                activebackground='#FFFF00',  # Brighter yellow on hover
                activeforeground='#004d00',  # Darker green on hover
                relief='raised',
                bd=2,
                font=('Arial', 14, 'bold'),
                width=20,
                height=2,
                cursor='hand2'
            )
            
            # Add hover effect
            def on_enter(e):
                start_button['bg'] = '#FFFF00'
            def on_leave(e):
                start_button['bg'] = '#FFD700'
                
            start_button.bind('<Enter>', on_enter)
            start_button.bind('<Leave>', on_leave)
            
            start_button.pack(pady=10)
            
            # Add a decorative border around the button
            border_frame = tk.Frame(
                button_frame,
                bg='#FFD700',
                padx=2,
                pady=2
            )
            border_frame.place(in_=start_button, relwidth=1.02, relheight=1.1,
                             relx=0.5, rely=0.5, anchor='center')
            start_button.lift()  # Ensure button stays on top of border
            
            def on_dialog_close():
                dialog.destroy()
            
            dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
            
        except Exception as e:
            logging.error(f"Dialog creation error: {e}")
            logging.error(traceback.format_exc())

    def start_cleanup(self, dialog, folders):
        """Start the cleanup process"""
        try:
            selected_folders = [f for f, v in folders.items() if v.get()]
            
            if not selected_folders:
                messagebox.showwarning(
                    "No Selection",
                    "Please select at least one folder to organize."
                )
                return
            
            dialog.destroy()
            
            def run_organizer():
                try:
                    organizer = FileOrganizer()
                    for folder in selected_folders:
                        if folder == 'Workspaces':
                            workspace_path = os.path.join(os.path.expanduser('~'), 'Workspaces')
                            if os.path.exists(workspace_path):
                                organizer.organize_folder(workspace_path)
                        else:
                            folder_path = organizer.get_special_folder_path(folder)
                            if folder_path:
                                organizer.organize_folder(folder_path)
                    
                    organizer.fix_onedrive_archives()
                    
                    # Show completion message
                    self.window.after(0, lambda: messagebox.showinfo(
                        "Complete",
                        "Folder organization complete!"
                    ))
                    
                except Exception as e:
                    logging.error(f"Cleanup error: {e}")
                    logging.error(traceback.format_exc())
                    self.window.after(0, lambda: messagebox.showerror(
                        "Error",
                        f"An error occurred during cleanup: {str(e)}"
                    ))
            
            threading.Thread(target=run_organizer, daemon=True).start()
            
        except Exception as e:
            logging.error(f"Cleanup start error: {e}")
            logging.error(traceback.format_exc())

    def show_menu(self, event):
        self.menu.tk_popup(event.x_root, event.y_root)

    def on_move(self, event):
        """Handle window dragging"""
        try:
            if hasattr(self, 'drag_start_x') and hasattr(self, 'drag_start_y'):
                # Calculate new position
                x = self.window.winfo_x() + (event.x_root - self.drag_start_x)
                y = self.window.winfo_y() + (event.y_root - self.drag_start_y)
                
                # Update drag start position
                self.drag_start_x = event.x_root
                self.drag_start_y = event.y_root
                
                # Move window
                self.window.geometry(f'+{x}+{y}')
        except Exception as e:
            logging.error(f"Move error: {e}")
            logging.error(traceback.format_exc())

    def start_move(self, event):
        """Initialize drag operation"""
        try:
            self.drag_start_x = event.x_root
            self.drag_start_y = event.y_root
        except Exception as e:
            logging.error(f"Start move error: {e}")
            logging.error(traceback.format_exc())

def create_startup_shortcut():
    if sys.platform == 'win32':
        import winreg
        import os
        
        # Get the path to the script
        script_path = os.path.abspath(sys.argv[0])
        
        # Create startup registry key
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, "DDAOrganizer", 0, winreg.REG_SZ, f'pythonw "{script_path}"')
            winreg.CloseKey(key)
        except Exception as e:
            logging.error(f"Failed to create startup entry: {e}")

def main():
    """Main entry point that hides console window"""
    try:
        if sys.platform == 'win32':
            # Hide console window on Windows
            try:
                import win32gui
                import win32con
                hwnd = win32gui.GetForegroundWindow()
                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            except Exception as e:
                logging.error(f"Failed to hide console: {e}")

        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('file_organizer.log'),
                logging.StreamHandler()
            ]
        )
        
        # Create startup shortcut
        create_startup_shortcut()
        
        # Start GUI
        app = FileOrganizerGUI()
        app.window.mainloop()
        
    except Exception as e:
        logging.error(f"Application error: {e}")
        logging.error(traceback.format_exc())
        
        # Show error message to user
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Error",
                f"Application failed to start: {str(e)}\nCheck file_organizer.log for details."
            )
        except:
            pass
        
        sys.exit(1)

if __name__ == "__main__":
    main()
