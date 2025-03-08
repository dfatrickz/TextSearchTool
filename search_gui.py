import os
import re
import time
from pathlib import Path
import fnmatch
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox

# Configuration settings
DIR = "/path/to/your/search/directory"
OUTPUT_DIR = "/path/to/your/output/directory"
DEFAULT_TERMS = ["apple"]
OUTPUT_FILE_TYPES = [".rtf", ".txt", ".md", ".docx", ".csv"]
DEFAULT_OUTPUT_FILE_TYPE = ".rtf"
EXCERPT_SENTENCES = 5
PROXIMITY_WINDOW = 5
IGNORE_STRING = r"(ignore these patterns)"
IGNORE_FILES = ["index.txt", "*.log"]
IGNORE_FOLDERS = ["temp", "logs"]
MIDDLE_SENTENCES = 10
MIDDLE_WORD_LIMIT = 150
UPDATE_INTERVAL = 50
RTF_HEADER = r"{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Calibri;}}{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}\f0\fs22\par"
RTF_FOOTER = r"}"

# Highlight styles mapping to RTF tags (only used for .rtf)
HIGHLIGHT_STYLES = {
    "Bold": r"\b {term}\b0",
    "Red": r"\cf1 {term}\cf0",
    "Blue": r"\cf2 {term}\cf0",
    "Bold Red": r"\b \cf1 {term}\cf0\b0",
    "Bold Blue": r"\b \cf2 {term}\cf0\b0"
}

# Search modes
SEARCH_MODES = ["Individual Mode", "Proximity Mode"]

# Global state
output_files = {}
start_time = 0
files_processed = 0
total_matches_by_term = {}
total_files = 0
running = False

class SearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text Search Tool")
        # Removed self.root.geometry("800x600") to allow dynamic resizing
        self.search_terms = tk.StringVar(value=",".join(DEFAULT_TERMS))
        self.search_dir = tk.StringVar(value=DIR)
        self.output_dir = tk.StringVar(value=OUTPUT_DIR)
        self.highlight_style = tk.StringVar(value="Bold")
        self.output_file_type = tk.StringVar(value=DEFAULT_OUTPUT_FILE_TYPE)
        self.search_mode = tk.StringVar(value=SEARCH_MODES[0])
        self.proximity_window = tk.StringVar(value=str(PROXIMITY_WINDOW))
        self.excerpt_sentences = tk.StringVar(value=str(EXCERPT_SENTENCES))
        self.middle_word_limit = tk.StringVar(value=str(MIDDLE_WORD_LIMIT))
        self.ignore_files = tk.StringVar(value=",".join(IGNORE_FILES))
        self.ignore_folders = tk.StringVar(value=",".join(IGNORE_FOLDERS))
        self.create_widgets()

    def create_widgets(self):
        # Configure root to expand
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Main frame for better layout
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure main_frame to expand
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)  # Row 5 (stats_text) will expand

        # Input Section
        input_frame = ttk.LabelFrame(main_frame, text="Search Inputs", padding="5")
        input_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        tk.Label(input_frame, text="Search Terms (comma-separated):").grid(row=0, column=0, pady=2, sticky=tk.W)
        tk.Entry(input_frame, textvariable=self.search_terms, width=50).grid(row=0, column=1, pady=2)

        tk.Label(input_frame, text="Search Directory:").grid(row=1, column=0, pady=2, sticky=tk.W)
        dir_frame = ttk.Frame(input_frame)
        dir_frame.grid(row=1, column=1, pady=2, sticky=(tk.W, tk.E))
        tk.Entry(dir_frame, textvariable=self.search_dir, width=40).pack(side=tk.LEFT)
        tk.Button(dir_frame, text="Browse", command=self.browse_search_dir).pack(side=tk.LEFT, padx=5)

        tk.Label(input_frame, text="Output Directory:").grid(row=2, column=0, pady=2, sticky=tk.W)
        out_frame = ttk.Frame(input_frame)
        out_frame.grid(row=2, column=1, pady=2, sticky=(tk.W, tk.E))
        tk.Entry(out_frame, textvariable=self.output_dir, width=40).pack(side=tk.LEFT)
        tk.Button(out_frame, text="Browse", command=self.browse_output_dir).pack(side=tk.LEFT, padx=5)

        # Mode and Customization Section
        mode_frame = ttk.LabelFrame(main_frame, text="Search Settings", padding="5")
        mode_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        tk.Label(mode_frame, text="Search Mode:").grid(row=0, column=0, pady=2, sticky=tk.W)
        mode_combo = ttk.Combobox(mode_frame, textvariable=self.search_mode, values=SEARCH_MODES, state="readonly")
        mode_combo.grid(row=0, column=1, pady=2, sticky=tk.W)
        mode_combo.set(SEARCH_MODES[0])

        tk.Label(mode_frame, text="Proximity Window (sentences, for Proximity Mode):").grid(row=1, column=0, pady=2, sticky=tk.W)
        proximity_frame = ttk.Frame(mode_frame)
        proximity_frame.grid(row=1, column=1, pady=2, sticky=(tk.W, tk.E))
        self.proximity_entry = tk.Entry(proximity_frame, textvariable=self.proximity_window, width=10)
        self.proximity_entry.pack(side=tk.LEFT)
        self.proximity_entry.config(state=tk.DISABLED)

        tk.Label(mode_frame, text="Excerpt Sentences Limit:").grid(row=2, column=0, pady=2, sticky=tk.W)
        ttk.Entry(mode_frame, textvariable=self.excerpt_sentences, width=10).grid(row=2, column=1, pady=2, sticky=tk.W)

        tk.Label(mode_frame, text="Middle Word Limit:").grid(row=3, column=0, pady=2, sticky=tk.W)
        ttk.Entry(mode_frame, textvariable=self.middle_word_limit, width=10).grid(row=3, column=1, pady=2, sticky=tk.W)

        # Ignore Settings Section
        ignore_frame = ttk.LabelFrame(main_frame, text="Ignore Settings", padding="5")
        ignore_frame.grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        tk.Label(ignore_frame, text="Ignore Files (comma-separated, e.g., index.txt, *.log):").grid(row=0, column=0, pady=2, sticky=tk.W)
        tk.Entry(ignore_frame, textvariable=self.ignore_files, width=50).grid(row=0, column=1, pady=2)

        tk.Label(ignore_frame, text="Ignore Folders (comma-separated, e.g., temp, logs):").grid(row=1, column=0, pady=2, sticky=tk.W)
        tk.Entry(ignore_frame, textvariable=self.ignore_folders, width=50).grid(row=1, column=1, pady=2)

        # Output Settings Section
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="5")
        output_frame.grid(row=3, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        tk.Label(output_frame, text="Highlight Style (for RTF only):").grid(row=0, column=0, pady=2, sticky=tk.W)
        style_combo = ttk.Combobox(output_frame, textvariable=self.highlight_style, values=list(HIGHLIGHT_STYLES.keys()), state="readonly")
        style_combo.grid(row=0, column=1, pady=2, sticky=tk.W)
        style_combo.set("Bold")

        tk.Label(output_frame, text="Output File Type:").grid(row=1, column=0, pady=2, sticky=tk.W)
        filetype_combo = ttk.Combobox(output_frame, textvariable=self.output_file_type, values=OUTPUT_FILE_TYPES, state="readonly")
        filetype_combo.grid(row=1, column=1, pady=2, sticky=tk.W)
        filetype_combo.set(DEFAULT_OUTPUT_FILE_TYPE)

        # Button Section
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, pady=10)
        self.start_btn = tk.Button(btn_frame, text="Start Search", command=self.start_search)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = tk.Button(btn_frame, text="Stop", command=self.stop_search, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # Stats Section
        stats_frame = ttk.LabelFrame(main_frame, text="Search Stats", padding="5")
        stats_frame.grid(row=5, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)

        # Start with minimal height, will adjust dynamically
        self.stats_text = scrolledtext.ScrolledText(stats_frame, height=1)
        self.stats_text.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Bind toggle for proximity input
        self.search_mode.trace("w", self.toggle_proximity_input)

        # Force the window to update its size after all widgets are added
        self.root.update_idletasks()

    def toggle_proximity_input(self, *args):
        if self.search_mode.get() == "Proximity Mode":
            self.proximity_entry.config(state=tk.NORMAL)
        else:
            self.proximity_entry.config(state=tk.DISABLED)

    def browse_search_dir(self):
        dir_path = filedialog.askdirectory(initialdir=self.search_dir.get())
        if dir_path:
            self.search_dir.set(dir_path)

    def browse_output_dir(self):
        dir_path = filedialog.askdirectory(initialdir=self.output_dir.get())
        if dir_path:
            self.output_dir.set(dir_path)

    def update_stats(self, terms):
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, f"Searching For: {' '.join(terms)}\n")
        self.stats_text.insert(tk.END, f"Files Processed: {files_processed}/{total_files}\n")
        if self.search_mode.get() == "Individual Mode":
            for term in terms:
                self.stats_text.insert(tk.END, f"Matches for {term}: {total_matches_by_term.get(term, 0)}\n")
        else:
            self.stats_text.insert(tk.END, f"Proximity Matches: {total_matches_by_term.get('proximity', 0)}\n")
        total_time = time.time() - start_time if start_time else 0
        speed = files_processed / total_time if total_time > 0 else 0
        self.stats_text.insert(tk.END, f"Speed (files/sec): {speed:.2f}\n")
        self.stats_text.insert(tk.END, f"Elapsed Time (sec): {total_time:.2f}\n")
        files_left = total_files - files_processed
        time_left = files_left / speed if speed > 0 else float('inf')
        self.stats_text.insert(tk.END, f"Est. Time Left (sec): {time_left:.2f}\n" if time_left != float('inf') else "N/A\n")

        # Dynamically adjust the height based on the number of lines
        content = self.stats_text.get(1.0, tk.END).strip()
        line_count = len(content.splitlines())
        # Set a reasonable maximum height to avoid excessive growth
        max_height = 30
        new_height = min(line_count + 1, max_height)  # Add 1 for padding
        self.stats_text.configure(height=new_height)

        # Force the window to resize based on content
        self.root.update_idletasks()
        self.root.update()

    def start_search(self):
        global running, start_time, files_processed, total_matches_by_term, total_files, output_files
        if running:
            return
        running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        files_processed = 0
        start_time = time.time()

        terms = [t.strip() for t in self.search_terms.get().split(',')] if self.search_terms.get() else DEFAULT_TERMS
        self.term_patterns = [re.compile(rf"(?:^|\s){term}(?=\s|$)", re.IGNORECASE) for term in terms]
        total_matches_by_term = {term: 0 for term in terms} if self.search_mode.get() == "Individual Mode" else {'proximity': 0}

        search_dir = self.search_dir.get()
        output_dir = self.output_dir.get()
        output_file_type = self.output_file_type.get()
        is_rtf = output_file_type == ".rtf"
        mode = self.search_mode.get()

        # Get configurable limits and store as instance variables
        try:
            self.excerpt_sentences_val = int(self.excerpt_sentences.get())
            self.middle_word_limit_val = int(self.middle_word_limit.get())
            self.proximity_window_val = int(self.proximity_window.get()) if mode == "Proximity Mode" else PROXIMITY_WINDOW
        except ValueError:
            self.excerpt_sentences_val = EXCERPT_SENTENCES
            self.middle_word_limit_val = MIDDLE_WORD_LIMIT
            self.proximity_window_val = PROXIMITY_WINDOW

        # Update ignore patterns from GUI
        self.ignore_files_list = [f.strip() for f in self.ignore_files.get().split(',')] if self.ignore_files.get() else IGNORE_FILES
        self.ignore_folders_list = [f.strip() for f in self.ignore_folders.get().split(',')] if self.ignore_folders.get() else IGNORE_FOLDERS

        if not os.path.isdir(search_dir):
            self.stats_text.insert(tk.END, f"Error: {search_dir} does not exist\n")
            self.stop_search()
            return

        os.makedirs(output_dir, exist_ok=True)

        # Setup output files based on mode
        overwrite_files = []
        if mode == "Individual Mode":
            for term in terms:
                output_file = os.path.join(output_dir, f"{term}{output_file_type}")
                if os.path.exists(output_file):
                    overwrite_files.append(f"{term}{output_file_type}")
                output_files[term] = open(output_file, "w", encoding="utf-8", newline="\n")
                if is_rtf:
                    output_files[term].write(RTF_HEADER)
        else:  # Proximity Mode
            if len(terms) < 2:
                self.stats_text.insert(tk.END, "Error: Proximity Mode requires at least 2 terms\n")
                self.stop_search()
                return
            output_file_name = "_".join(terms).lower() + output_file_type
            output_file = os.path.join(output_dir, output_file_name)
            if os.path.exists(output_file):
                overwrite_files.append(output_file_name)
            output_files['proximity'] = open(output_file, "w", encoding="utf-8", newline="\n")
            if is_rtf:
                output_files['proximity'].write(RTF_HEADER)

        # Overwrite warning
        if overwrite_files:
            warning_msg = "Warning: You're about to overwrite the following existing output files:\n" + "\n".join(overwrite_files) + "\n\nContinue?"
            if not messagebox.askyesno("Overwrite Warning", warning_msg):
                self.stop_search()
                return
        self.txt_files = [f for f in Path(search_dir).rglob("*.[tT][xX][tT]")
                         if not (any(fnmatch.fnmatch(f.name, ignore) for ignore in self.ignore_files_list) or
                                 any(fnmatch.fnmatch(str(f.parent.name), ignore) for ignore in self.ignore_folders_list))]
        total_files = len(self.txt_files)
        self.update_stats(terms)

        # Start the search loop
        self.root.after(10, self.search_loop)

    def search_loop(self):
        global files_processed
        ignore_pattern = re.compile(IGNORE_STRING)
        mode = self.search_mode.get()
        is_rtf = self.output_file_type.get() == ".rtf"
        terms = [t.strip() for t in self.search_terms.get().split(',')] if self.search_terms.get() else DEFAULT_TERMS

        for file in self.txt_files:
            if not running:
                break
            file_start = time.time()
            with open(file, "r", encoding="utf-8", errors="ignore") as f:
                raw_text = f.read().replace(r'\c', r'\\c')
                lines = raw_text.splitlines()
                sentences_all = re.split(r'\.\s*', raw_text)
                sentences_all = [s.strip() + "." for s in sentences_all if s.strip()]

            if mode == "Individual Mode":
                for i, line in enumerate(lines, 1):
                    if ignore_pattern.search(line):
                        continue
                    for term_idx, pattern in enumerate(self.term_patterns):
                        term = terms[term_idx]
                        match = pattern.search(line)
                        if match:
                            total_matches_by_term[term] += 1
                            start = max(0, i - self.proximity_window_val)
                            end = min(len(lines), i + self.proximity_window_val)
                            excerpt_lines = [l for l in lines[start:end] if not ignore_pattern.search(l)]
                            excerpt_full = " ".join(excerpt_lines)
                            sentences = re.split(r'\.\s*', excerpt_full)
                            keyword_sentence_idx = -1
                            for idx, sentence in enumerate(sentences):
                                if pattern.search(sentence):
                                    keyword_sentence_idx = idx
                                    break
                            if keyword_sentence_idx == -1:
                                keyword_excerpt = line.strip() + "."
                            else:
                                excerpt_start = max(0, keyword_sentence_idx - (self.excerpt_sentences_val // 2))
                                excerpt_end = min(len(sentences), excerpt_start + self.excerpt_sentences_val)
                                keyword_excerpt = " ".join(sentences[excerpt_start:excerpt_end]).strip() + "."
                                if not pattern.search(keyword_excerpt):
                                    keyword_excerpt = line.strip() + "."
                            matched_term = match.group(0)
                            if is_rtf:
                                keyword_excerpt = keyword_excerpt.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')
                                style = self.highlight_style.get()
                                highlighted_term = HIGHLIGHT_STYLES[style].format(term=matched_term)
                                keyword_excerpt = keyword_excerpt.replace(matched_term, highlighted_term)
                            else:
                                keyword_excerpt = keyword_excerpt

                            mid_point = len(sentences_all) // 2
                            mid_start = max(0, mid_point - (MIDDLE_SENTENCES // 2))
                            mid_end = min(len(sentences_all), mid_start + MIDDLE_SENTENCES)
                            middle_sentences = sentences_all[mid_start:mid_end]
                            middle_excerpt = " ".join(middle_sentences)
                            words = middle_excerpt.split()
                            if len(words) > self.middle_word_limit_val:
                                middle_excerpt = " ".join(words[:self.middle_word_limit_val]) + "..."
                            if is_rtf:
                                middle_excerpt = middle_excerpt.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')

                            out_f = output_files[term]
                            out_f.write(f"File: {file}\n")
                            out_f.write(f"Match at line {i} (keyword excerpt): {keyword_excerpt}\n")
                            out_f.write("\n")
                            out_f.write("\n")
                            out_f.write(f"Middle of file excerpt: {middle_excerpt}\n")
                            out_f.write("------------------------\n" if not is_rtf else "------------------------\\par\n")
                            if is_rtf:
                                out_f.write("\\par\n")
                            out_f.flush()
            else:
                proximity_matches = []
                for i, sentence in enumerate(sentences_all):
                    if ignore_pattern.search(sentence):
                        continue
                    if self.term_patterns[0].search(sentence):
                        start = max(0, i - self.proximity_window_val)
                        end = min(len(sentences_all), i + self.proximity_window_val + 1)
                        window = " ".join(sentences_all[start:end])
                        if all(pattern.search(window) for pattern in self.term_patterns):
                            proximity_matches.append(i)

                for match_idx in proximity_matches:
                    total_matches_by_term['proximity'] += 1
                    start = max(0, match_idx - (self.excerpt_sentences_val // 2))
                    end = min(len(sentences_all), start + self.excerpt_sentences_val)
                    excerpt_sentences_list = sentences_all[start:end]
                    keyword_excerpt = " ".join(excerpt_sentences_list)

                    if not all(pattern.search(keyword_excerpt) for pattern in self.term_patterns):
                        continue

                    if is_rtf:
                        keyword_excerpt = keyword_excerpt.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')
                        style = self.highlight_style.get()
                        for pattern_idx, pattern in enumerate(self.term_patterns):
                            matches = list(pattern.finditer(keyword_excerpt))
                            for match in reversed(matches):
                                matched_term = match.group(0)
                                highlighted_term = HIGHLIGHT_STYLES[style].format(term=matched_term)
                                keyword_excerpt = (keyword_excerpt[:match.start()] + 
                                                 highlighted_term + 
                                                 keyword_excerpt[match.end():])
                    else:
                        keyword_excerpt = keyword_excerpt

                    mid_point = len(sentences_all) // 2
                    mid_start = max(0, mid_point - (MIDDLE_SENTENCES // 2))
                    mid_end = min(len(sentences_all), mid_start + MIDDLE_SENTENCES)
                    middle_sentences = sentences_all[mid_start:mid_end]
                    middle_excerpt = " ".join(middle_sentences)
                    words = middle_excerpt.split()
                    if len(words) > self.middle_word_limit_val:
                        middle_excerpt = " ".join(words[:self.middle_word_limit_val]) + "..."
                    if is_rtf:
                        middle_excerpt = middle_excerpt.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')

                    out_f = output_files['proximity']
                    out_f.write(f"File: {file}\n")
                    out_f.write(f"Proximity match at sentence {match_idx + 1} (keyword excerpt): {keyword_excerpt}\n")
                    out_f.write("\n")
                    out_f.write("\n")
                    out_f.write(f"Middle of file excerpt: {middle_excerpt}\n")
                    out_f.write("------------------------\n" if not is_rtf else "------------------------\\par\n")
                    if is_rtf:
                        out_f.write("\\par\n")
                    out_f.flush()

            files_processed += 1
            if files_processed % UPDATE_INTERVAL == 0:
                elapsed = time.time() - file_start
                self.stats_text.insert(tk.END, f"Processed {file} in {elapsed:.2f} seconds\n")
                self.update_stats(terms)
                self.root.update()
                if not running:
                    break

        self.finalize_search(terms)

    def stop_search(self):
        global running
        if running:
            running = False
            self.stats_text.insert(tk.END, "Stopping search...\n")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def finalize_search(self, terms):
        global output_files, running
        for key, f in output_files.items():
            if f and not f.closed:
                if self.output_file_type.get() == ".rtf":
                    f.write(RTF_FOOTER)
                f.close()
        self.update_stats(terms)
        if running:
            self.stats_text.insert(tk.END, "\nSearch Completed Successfully\n")
            running = False
        else:
            self.stats_text.insert(tk.END, "\nSearch Stopped by User\n")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = SearchApp(root)
    root.mainloop()
