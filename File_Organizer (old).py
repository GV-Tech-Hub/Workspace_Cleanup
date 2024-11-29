import os
import logging
import json
import platform
import shutil
import time
import traceback
import datetime
import ctypes
from pathlib import Path
from typing import Optional
import winreg

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
        """Get the path of special folders with multi-OS support and multiple path checking"""
        try:
            if self.os_type == "Windows":
                possible_paths = []
                
                # Add OneDrive path if it exists
                onedrive_path = os.path.join(str(Path.home()), "OneDrive", folder_name)
                if os.path.exists(onedrive_path):
                    possible_paths.append(onedrive_path)
                    logging.info(f"Found OneDrive {folder_name} path: {onedrive_path}")
                
                # Add regular Windows path
                regular_path = os.path.join(str(Path.home()), folder_name)
                if os.path.exists(regular_path):
                    possible_paths.append(regular_path)
                    logging.info(f"Found regular {folder_name} path: {regular_path}")
                
                # For Downloads, also check Windows Registry
                if folder_name == "Downloads":
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                          r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                            reg_path = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
                            if os.path.exists(reg_path):
                                possible_paths.append(reg_path)
                                logging.info(f"Found Registry {folder_name} path: {reg_path}")
                    except Exception as e:
                        logging.warning(f"Failed to get Downloads path from registry: {e}")
                
                # If we found any valid paths, use the first one
                if possible_paths:
                    chosen_path = possible_paths[0]
                    logging.info(f"Using {folder_name} path: {chosen_path}")
                    return chosen_path
                else:
                    logging.error(f"No valid {folder_name} path found")
                    return None
                
            else:  # macOS and Linux
                path = str(Path.home() / folder_name)
                if os.path.exists(path):
                    logging.info(f"Using {self.os_type} {folder_name} path: {path}")
                    return path
                else:
                    logging.error(f"No valid {folder_name} path found for {self.os_type}")
                    return None
                
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

    def organize_folder(self, folder_path: str) -> None:
        """Organize files with enhanced archive handling"""
        try:
            if not folder_path or not os.path.exists(folder_path):
                logging.error(f"Invalid or non-existent folder path: {folder_path}")
                return

            # Create or get the root archive folder (always named "Archive")
            archive_folder = os.path.join(folder_path, "Archive")
            if not os.path.exists(archive_folder):
                os.makedirs(archive_folder)
                logging.info(f"Created archive folder: {archive_folder}")

            # Create new dated subfolder
            current_time = datetime.datetime.now()
            timestamp = current_time.strftime("%b-%d-%Y_%I-%M%p")
            dated_subfolder = os.path.join(archive_folder, timestamp)
            
            # Get all items in the directory
            try:
                all_items = os.listdir(folder_path)
            except Exception as e:
                logging.error(f"Failed to list directory contents for {folder_path}: {e}")
                return

            # Filter items to move
            files_to_move = []
            for item in all_items:
                item_path = os.path.join(folder_path, item)
                
                # Skip if:
                # 1. It's the Archive folder itself
                # 2. It's a system file or Recycle Bin
                # 3. It's any kind of archive folder
                if (item != "Archive" and 
                    not self.is_system_file(item_path) and
                    not "archive" in item.lower()):
                    files_to_move.append(item)

            if files_to_move:
                # Create the dated subfolder
                os.makedirs(dated_subfolder, exist_ok=True)
                logging.info(f"Created dated subfolder: {timestamp}")
                
                for item in files_to_move:
                    try:
                        item_path = os.path.join(folder_path, item)
                        target_path = os.path.join(dated_subfolder, item)
                        
                        # Handle file in use errors
                        retry_count = 3
                        while retry_count > 0:
                            try:
                                # Update statistics before moving
                                self.stats["space_cleared"] += self.get_file_size(item_path)
                                shutil.move(item_path, target_path)
                                self.stats["files_moved"] += 1
                                logging.info(f"Moved: {item}")
                                break
                            except PermissionError:
                                retry_count -= 1
                                time.sleep(1)
                                if retry_count == 0:
                                    raise
                        
                    except Exception as e:
                        self.stats["errors"] += 1
                        logging.error(f"Error moving {item}: {e}")
            
            else:
                logging.info(f"No files to organize in {folder_path}")

        except Exception as e:
            logging.error(f"Organization failed for {folder_path}: {e}")
            logging.error(traceback.format_exc())

    def organize(self):
        """Main organization method"""
        try:
            logging.info("Starting organization process...")

            if self.desktop_path:
                logging.info("Organizing Desktop...")
                self.organize_folder(self.desktop_path)
            else:
                logging.error("Skipping Desktop organization - path not available")

            if self.downloads_path:
                logging.info("Organizing Downloads...")
                self.organize_folder(self.downloads_path)
            else:
                logging.error("Skipping Downloads organization - path not available")

            self.print_statistics()

        except Exception as e:
            logging.error(f"Organization failed: {e}")
            logging.error(traceback.format_exc())
            self.stats["errors"] += 1
        finally:
            logging.info("Organization complete!")

    def print_statistics(self):
        """Print organization statistics"""
        logging.info("\n=== Organization Statistics ===")
        logging.info(f"Files moved: {self.stats['files_moved']}")
        logging.info(f"Space cleared: {self.stats['space_cleared'] / (1024*1024):.2f} MB")
        logging.info(f"Errors encountered: {self.stats['errors']}")
        logging.info("===========================")

def main():
    try:
        print("Starting File Organizer...")
        organizer = FileOrganizer()
        print("Initialization successful!")
        
        user_input = input("Press Enter to start organizing files (or 'q' to quit): ")
        if user_input.lower() != 'q':
            organizer.organize()
            print("\nOrganization complete!")
        
        input("Press Enter to exit...")
    
    except Exception as e:
        print(f"\nError occurred: {str(e)}")
        logging.error(f"Error occurred: {str(e)}")
        logging.error(traceback.format_exc())
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()