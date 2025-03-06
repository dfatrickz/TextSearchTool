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
OUTPUT_FILE_TYPES = [".rtf", ".txt", ".md", ".docx", ".csv"]  # Common text file types
DEFAULT_OUTPUT_FILE_TYPE = ".rtf"  # Default file type
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
        self.root.geometry("800x600")
        self.search_terms = tk.StringVar(value=",".join(DEFAULT_TERMS))
        self.search_dir = tk.StringVar(value=DIR)
        self.output_dir = tk.StringVar(value=OUTPUT_DIR)
        self.highlight_style = tk.StringVar(value="Bold")  # Default style for RTF
        self.output_file_type = tk.StringVar(value=DEFAULT_OUTPUT_FILE_TYPE)  # Default file type
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Search Terms (comma-separated):").pack(pady=5)
        tk.Entry(self.root, textvariable=self.search_terms, width=50).pack()
        
        tk.Label(self.root, text="Search Directory:").pack(pady=5)
        dir_frame = tk.Frame(self.root)
        dir_frame.pack()
        tk.Entry(dir_frame, textvariable=self.search_dir, width=40).pack(side=tk.LEFT)
        tk.Button(dir_frame, text="Browse", command=self.browse_search_dir).pack(side=tk.LEFT, padx=5)
        
        tk.Label(self.root, text="Output Directory:").pack(pady=5)
        out_frame = tk.Frame(self.root)
        out_frame.pack()
        tk.Entry(out_frame, textvariable=self.output_dir, width=40).pack(side=tk.LEFT)
        tk.Button(out_frame, text="Browse", command=self.browse_output_dir).pack(side=tk.LEFT, padx=5)
        
        # Highlight style selection (for RTF only)
        tk.Label(self.root, text="Highlight Style (for RTF only):").pack(pady=5)
        style_frame = tk.Frame(self.root)
        style_frame.pack()
        style_combo = ttk.Combobox(style_frame, textvariable=self.highlight_style, 
                                 values=list(HIGHLIGHT_STYLES.keys()), state="readonly")
        style_combo.pack()
        style_combo.set("Bold")  # Default selection
        
        # Output file type selection
        tk.Label(self.root, text="Output File Type:").pack(pady=5)
        filetype_frame = tk.Frame(self.root)
        filetype_frame.pack()
        filetype_combo = ttk.Combobox(filetype_frame, textvariable=self.output_file_type, 
                                    values=OUTPUT_FILE_TYPES, state="readonly")
        filetype_combo.pack()
        filetype_combo.set(DEFAULT_OUTPUT_FILE_TYPE)  # Default selection
        
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.start_btn = tk.Button(btn_frame, text="Start Search", command=self.start_search)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = tk.Button(btn_frame, text="Stop", command=self.stop_search, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.stats_text = scrolledtext.ScrolledText(self.root, height=20, width=80)
        self.stats_text.pack(pady=10)

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
        for term in terms:
            self.stats_text.insert(tk.END, f"Matches for {term}: {total_matches_by_term.get(term, 0)}\n")
        total_time = time.time() - start_time if start_time else 0
        speed = files_processed / total_time if total_time > 0 else 0
        self.stats_text.insert(tk.END, f"Speed (files/sec): {speed:.2f}\n")
        self.stats_text.insert(tk.END, f"Elapsed Time (sec): {total_time:.2f}\n")
        files_left = total_files - files_processed
        time_left = files_left / speed if speed > 0 else float('inf')
        self.stats_text.insert(tk.END, f"Est. Time Left (sec): {time_left:.2f}\n" if time_left != float('inf') else "N/A\n")
        self.root.update_idletasks()

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
        term_patterns = {term: re.compile(rf"(?:^|\s){term}(?=\s|$)", re.IGNORECASE) for term in terms}
        total_matches_by_term = {term: 0 for term in terms}

        search_dir = self.search_dir.get()
        output_dir = self.output_dir.get()
        output_file_type = self.output_file_type.get()  # Get selected file type
        is_rtf = output_file_type == ".rtf"  # Check if RTF for formatting

        if not os.path.isdir(search_dir):
            self.stats_text.insert(tk.END, f"Error: {search_dir} does not exist\n")
            self.stop_search()
            return

        os.makedirs(output_dir, exist_ok=True)
        
        # Check for existing files and warn user
        overwrite_files = []
        for term in terms:
            output_file = os.path.join(output_dir, f"{term}{output_file_type}")
            if os.path.exists(output_file):
                overwrite_files.append(f"{term}{output_file_type}")
        
        if overwrite_files:
            warning_msg = "Warning: You're about to overwrite the following existing output files:\n" + "\n".join(overwrite_files) + "\n\nContinue?"
            if not messagebox.askyesno("Overwrite Warning", warning_msg):
                self.stop_search()
                return

        # Open output files with appropriate headers
        for term in terms:
            output_file = os.path.join(output_dir, f"{term}{output_file_type}")
            output_files[term] = open(output_file, "w", encoding="utf-8", newline="\n")
            if is_rtf:
                output_files[term].write(RTF_HEADER)

        txt_files = [f for f in Path(search_dir).rglob("*.[tT][xX][tT]")
                     if not (any(fnmatch.fnmatch(f.name, ignore) for ignore in IGNORE_FILES) or
                             any(fnmatch.fnmatch(str(f.parent.name), ignore) for ignore in IGNORE_FOLDERS))]
        total_files = len(txt_files)
        self.update_stats(terms)

        def search_loop():
            global files_processed
            ignore_pattern = re.compile(IGNORE_STRING)
            for file in txt_files:
                if not running:
                    break
                file_start = time.time()
                with open(file, "r", encoding="utf-8", errors="ignore") as f:
                    raw_text = f.read().replace(r'\c', r'\\c')
                    lines = raw_text.splitlines()
                    sentences_all = re.split(r'\.\s*', raw_text)
                    sentences_all = [s.strip() + "." for s in sentences_all if s.strip()]

                for i, line in enumerate(lines, 1):
                    if ignore_pattern.search(line):
                        continue
                    for term, pattern in term_patterns.items():
                        match = pattern.search(line)
                        if match:
                            total_matches_by_term[term] += 1
                            # Keyword excerpt
                            start = max(0, i - PROXIMITY_WINDOW)
                            end = min(len(lines), i + PROXIMITY_WINDOW)
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
                                excerpt_start = max(0, keyword_sentence_idx - (EXCERPT_SENTENCES // 2))
                                excerpt_end = min(len(sentences), excerpt_start + EXCERPT_SENTENCES)
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
                                keyword_excerpt = keyword_excerpt  # Plain text, no RTF formatting

                            # Middle excerpt
                            mid_point = len(sentences_all) // 2
                            mid_start = max(0, mid_point - (MIDDLE_SENTENCES // 2))
                            mid_end = min(len(sentences_all), mid_start + MIDDLE_SENTENCES)
                            middle_sentences = sentences_all[mid_start:mid_end]
                            middle_excerpt = " ".join(middle_sentences)
                            words = middle_excerpt.split()
                            if len(words) > MIDDLE_WORD_LIMIT:
                                middle_excerpt = " ".join(words[:MIDDLE_WORD_LIMIT]) + "..."
                            if is_rtf:
                                middle_excerpt = middle_excerpt.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')

                            # Write to file
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

                files_processed += 1
                if files_processed % UPDATE_INTERVAL == 0:
                    elapsed = time.time() - file_start
                    self.stats_text.insert(tk.END, f"Processed {file} in {elapsed:.2f} seconds\n")
                    self.update_stats(terms)
                    self.root.update()
                    if not running:
                        break

            self.finalize_search(terms)

        self.root.after(10, search_loop)

    def stop_search(self):
        global running
        if running:
            running = False
            self.stats_text.insert(tk.END, "Stopping search...\n")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def finalize_search(self, terms):
        global output_files, running
        for term, f in output_files.items():
            if f and not f.closed:
                if self.output_file_type.get() == ".rtf":  # Only add footer for RTF
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
