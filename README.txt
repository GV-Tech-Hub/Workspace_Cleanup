 
 
# DesktopDownloadsArchiver (DDA)

## Why This Tool?
Created to solve the common problem of cluttered Desktop and Downloads folders. Instead of manually organizing or deleting files, DDA automatically moves them into dated archive folders, making it easy to:
- Keep workspaces clean
- Track when files were archived
- Maintain file history
- Quickly find old files by date

## How It Works
Built in Python, DDA:
1. Creates an "Archive" folder in Desktop & Downloads
2. Makes dated subfolders (e.g., "Nov-24-2023_02-30PM")
3. Moves all non-system files into these dated folders
4. Works with both OneDrive and local folders
5. Maintains detailed logs of all operations

## Usage
1. Run the script
2. Press Enter to start organizing
3. Check the new "Archive" folders in Desktop & Downloads

## Example Use Case
John's desktop has accumulated 3 months of files:
- Screenshots from meetings
- Downloaded PDFs
- Temporary files
Instead of spending time manually organizing or risking deleting important files, John runs DDA and instantly has everything organized by date in archive folders.

## Installation

### Requirements


bash
Create a virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate # On Windows: venv\Scripts\activate
Install required packages
pip install pathlib
pip install typing



### Setup
1. Download `DesktopDownloadsArchiver.py` and `config.json`
2. Place them in the same folder
3. Run `python DesktopDownloadsArchiver.py`

### Compatibility
- Windows ✅
- macOS ✅
- Linux ✅
- OneDrive Support ✅

## Configuration
Edit `config.json` to:
- Exclude specific file types
- Add system files to ignore
- Change archive folder name
- Set maximum archive age



Installation Commands (Copy/Paste Ready):

python -m venv venv
venv\Scripts\activate
pip install pathlib
pip install typing
python DesktopDownloadsArchiver.py