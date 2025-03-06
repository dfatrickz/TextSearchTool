# Text Search Tool

A Python-based GUI application designed to search for specific terms within text files in a specified directory and generate detailed output files with excerpts of matches and context. This tool is ideal for users who need to quickly analyze text data, such as researchers, developers, or anyone working with large sets of text files.

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
- Pattern Ignoring: Skips lines matching a configurable ignore pattern (default: (ignore these patterns)).
- Overwrite Protection: Warns users before overwriting existing output files with a confirmation dialog.
- Stop Functionality: Allows interrupting the search process mid-execution.
- Customizable Ignore Settings: Input file patterns (e.g., index.txt, *.log) and folder patterns (e.g., temp, logs) to exclude from search via GUI.
- Adjustable Excerpt Limits: Customize the number of sentences in keyword excerpts and the word limit for middle excerpts.
- Enhanced GUI Layout: Organized into sections using frames and grid layout for a cleaner, more user-friendly interface.

## Requirements

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
   python search_gui.py

Usage

1. Launch the tool with python search_gui.py.
2. Enter search terms (e.g., apple, banana) in the "Search Terms" field.
3. Select a search directory containing .txt files using the "Browse" button.
4. Choose an output directory for results.
5. (Optional) Select an output file type (default: .rtf) and, for .rtf, a highlight style.
6. Click "Start Search" to begin.
7. Monitor progress in the text area; click "Stop" to halt if needed.
8. Check the output directory for files named after each search term (e.g., apple.rtf).

Notes
- File Support: Currently searches only .txt files (case-insensitive, e.g., .TXT, .txt).
- RTF Formatting: Highlighting and special formatting apply only to .rtf output; other formats use plain text.
- Future Enhancements: Potential support for full .docx formatting or .csv structuring with additional libraries.

Contributions and feedback are welcome!







