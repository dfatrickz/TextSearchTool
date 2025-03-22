# Text Search Tool

A Privacy focused Python-based GUI application designed to search for specific terms within text files in a specified directory and generate detailed output files with excerpts of matches and context. This tool is ideal for users who need to quickly analyze text data, such as researchers, developers, or anyone working with large sets of text files.

## Recent Updates - March 22, 2025
- Added `.docx` output support with keyword highlighting (bold, red, blue).
- Improved `.docx` performance: batches save every 1000 matches, merging into one file per keyword at the end.
- Fixed bugs and optimized for stability.


## Speed Test (SATA SSD, Single-Keyword, Default Settings)

- Average file size: 33.15 KB
- Average character count: 33243.44 characters

- Searching For: Happy
- Files Processed: 62123/62123
- Matches for Happy: 62689
- Speed (files/sec): 1259.18
- Elapsed Time (sec): 53.04

## Features

Features

- Graphical User Interface (GUI): Built with tkinter, providing an intuitive interface for ease of use.
- Customizable Search Terms: Enter multiple search terms (comma-separated) to find matches across text files.
- Search Mode Selection: Choose between:
  - Individual Mode: Search for each term independently, outputting matches to separate files per term.
  - Proximity Mode: Search for all terms within a configurable proximity window (e.g., 10 sentences), outputting matches to a single file.
- Directory Selection: Browse and select both the search directory (where text files are located) and the output directory (where results are saved).
- Multiple Output Formats: Choose from common text file types for output:
  - .rtf (Rich Text Format, with highlighting support)
  - .txt (Plain text)
  - .md (Markdown)
  - .docx (Plain text with .docx extension)
  - .csv (Plain text with .csv extension)
- Highlighting for RTF: Customize how matched terms appear in .rtf files with options like Bold, Red, Blue, Bold Red, or Bold Blue.
- Excerpt Generation:
  - Keyword Excerpt: Displays the line containing the match plus a configurable number of surrounding lines (default: 5 sentences).
  - Middle Excerpt: Includes a snippet from the middle of the file (default: 10 sentences, up to 150 words) for context.
- Progress Tracking: Real-time stats in the GUI, including:
  - Files processed vs. total files
  - Matches per term (Individual Mode) or total proximity matches (Proximity Mode)
  - Search speed (files per second)
  - Elapsed time and estimated time remaining
- File and Folder Filtering: Ignores specified files (e.g., index.txt, *.log) and folders (e.g., temp, logs) to focus on relevant content.
- Pattern Ignoring: Skips lines matching a configurable ignore pattern (default: (ignore these patterns)). (Not in GUI yet)
- Overwrite Protection: Warns users before overwriting existing output files with a confirmation dialog.
- Stop Functionality: Allows interrupting the search process mid-execution.
- Customizable Ignore Settings: Input file patterns (e.g., index.txt, *.log) and folder patterns (e.g., temp, logs) to exclude from search via GUI.
- Adjustable Excerpt Limits: Customize the number of sentences in keyword excerpts and the word limit for middle excerpts.
- Enhanced GUI Layout: Organized into sections using frames and grid layout for a cleaner, more user-friendly interface.

## Requirements
- Python 3.x
- `tkinter` (included with Python)
- `python-docx` (install via `pip install python-docx` for .docx support)

To run the Text Search Tool, ensure you have the following:

### Software
- **Python 3.x**: The tool is written in Python 3 and requires a compatible interpreter.
  - Tested with Python 3.8+, but should work with most 3.x versions.
- **Operating System**: Compatible with Windows, macOS, or Linux (wherever Python and `tkinter` are supported).

### Python Libraries
The tool uses standard Python libraries, most of which come pre-installed with Python. Ensure these are available:
- `os`: For file and directory operations.
- `re`: For regular expression-based search term matching.
- `time`: For performance tracking.
- `pathlib`: For modern file path handling.
- `fnmatch`: For file and folder pattern matching.
- `tkinter`: For the GUI (usually included with Python; see installation note below if missing).

#### Installing `tkinter` (if not included)
- On **Ubuntu/Debian**:
  ```bash
  sudo apt-get install python3-tk

- On Fedora:
  sudo dnf install python3-tkinter
- On macOS/Windows: Typically bundled with Python; if not, reinstall Python from python.org with the tcl/tk option enabled.

No additional external packages (e.g., via pip install) are required beyond the standard library.

Installation

1. Clone the Repository:
   git clone https://github.com/dfatrickz/TextSearchTool.git
   cd TextSearchTool

2. Run the Tool:
   python3 search_gui.py
   
## Running the Script on Windows

Follow these steps to run `search_gui.py` on Windows:

1. **Install Python**:
   - Download Python 3.x from [python.org](https://www.python.org/downloads/).
   - Run the installer. **Check "Add Python to PATH"** during setup (bottom of the first screen).
   - Verify installation: Open Command Prompt (`cmd`) and type `python --version`. You should see something like `Python 3.10.0`.

2. **Install `python-docx` (Optional)**:
   - If you want `.docx` output support, open Command Prompt and run:
     ```
     pip install python-docx
     ```
   - If `pip` isn’t recognized, use:
     ```
     python -m pip install python-docx
     ```

3. **Run the Script**:
   - **Option 1: Double-Click**:
     - Save `search_gui.py` to a folder.
     - Double-click the file. If Python is installed correctly, the GUI should open.
   - **Option 2: Command Prompt**:
     - Open Command Prompt.
     - Navigate to the script’s folder:
       ```
       cd path\to\the\ScriptFolder
       ```
     - Run:
       ```
       python search_gui.py
       ```
   - If you see an error like `'python' is not recognized`, ensure Python was added to PATH (reinstall with the checkbox enabled).

4. **Troubleshooting**:
   - If `.docx` output fails with an error about `python-docx`, install it (step 2).
   - If the GUI doesn’t appear, ensure `tkinter` is installed by running `python -m tkinter` in Command Prompt—it should open a small test window. tkinter is included with the official Python installer, but you must select it during installation (it’s on by default in recent versions)




## Usage

1. Launch the tool with python3 search_gui.py.
2. Enter search terms (e.g., apple, banana) in the "Search Terms" field.
3. Select a search directory containing .txt files using the "Browse" button.
4. Choose an output directory for results.
5. (Optional) Select an output file type (default: .rtf) and, for .rtf/.docx, a highlight style.
6. Click "Start Search" to begin.
7. Monitor progress in the text area; click "Stop" to halt if needed.
8. Check the output directory for files named after each search term (e.g., apple.rtf).

Notes
- File Support: Currently searches for .txt files (case-insensitive, e.g., .TXT, .txt). and text files without ".txt"
- RTF Formatting: Highlighting and special formatting apply only to .rtf and .docx output; other formats use plain text.
- Future Enhancements: .csv structuring with additional libraries. (.csv has been removed for now)

Contributions and feedback are welcome!







